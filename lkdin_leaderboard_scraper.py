"""
LinkedIn Games Leaderboard Scraper
Automates browser to fetch leaderboard data from LinkedIn games and exports to Excel.

Fix changelog:
- Migrated from Selenium + webdriver_manager to Playwright (async)
  to match the proven pattern from naukri_updater.py (Arch Linux, Wayland, Brave)
- Uses launch_persistent_context with /usr/bin/brave + a local brave_profile dir
- Pinpoint scraped first; script exits early if it fails (confirms login + page structure is working)
- Results page scraped first for user score + average
- Average extracted from .pr-golden-chiclet__subtext containing "Today's avg:"
- User score extracted from .pr-golden-chiclet__text (the big number/time shown on results page)
- Leaderboard scrape targets .pr-connections-leaderboard__content for ranked players
- Fixed "You" detection: name is in .pr-connections-leaderboard-player__content-column .text-body-medium-bold
- Fixed duplicate-row bug: LinkedIn renders the "You" row twice on the
  leaderboard page (pinned + at actual rank), so the scraper could collect
  two "You" entries in one run. pd.merge() with a non-unique key does a
  Cartesian match, so those duplicates multiplied every day (1x2=2, 2x2=4,
  4x4=16, ...). Now: (1) "You" rows are collapsed to one right after
  scraping, (2) a general per-name dedupe runs as a safety net, and
  (3) export_to_excel dedupes both the freshly-scraped data and the
  existing sheet before merging, via dedupe_by_player_name().
"""

import asyncio
import re
import argparse
import os
import sys
import datetime
import pandas as pd
from playwright.async_api import async_playwright, TimeoutError as PlaywrightTimeoutError
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# Game order: pinpoint MUST be first so we can bail early on failure
# ---------------------------------------------------------------------------
GAMES = {
    "pinpoint": "Pinpoint",
    "wend": "Wend",
    "mini-sudoku": "Mini Sudoku",
    "zip": "Zip",
    "crossclimb": "Crossclimb",
    "queens": "Queens",
    "tango": "Tango",
    "patches": "Patches",
}

RESULTS_URL = "https://www.linkedin.com/games/{game}/results/"
LEADERBOARD_URL = (
    "https://www.linkedin.com/games/{game}/results/leaderboard/connections/"
)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def clean_score(text: str) -> str:
    """Strip leading apostrophes/backticks that LinkedIn sometimes injects."""
    if not text:
        return ""
    return re.sub(r"^['`\u2019]+", "", text.strip()).strip()


def dedupe_by_player_name(df: pd.DataFrame) -> pd.DataFrame:
    """Collapse rows that share a 'Player Name' into a single row.

    Uses groupby().first(), which keeps the first *non-null* value in each
    date column across the duplicate rows, rather than naively dropping
    duplicates and risking loss of data from a row whose duplicate had a
    different column populated.
    """
    if df.empty or "Player Name" not in df.columns:
        return df
    return df.groupby("Player Name", as_index=False, sort=False).first()


def parse_time_seconds(score: str) -> float:
    """Convert M:SS to float seconds for sorting. Returns large number on failure."""
    m = re.match(r"^(\d+):(\d{2})$", score.strip())
    if m:
        return int(m.group(1)) * 60 + int(m.group(2))
    m2 = re.match(r"^(\d+)$", score.strip())
    if m2:
        return float(m2.group(1))
    return 1e9


# ---------------------------------------------------------------------------
# Main scraper class
# ---------------------------------------------------------------------------


class LinkedInLeaderboardScraper:
    def __init__(self, browser_executable_path=None, headless=True):
        self.browser_executable_path = browser_executable_path or "/usr/bin/brave"
        self.headless = headless
        self.browser = None
        self.page = None
        self.all_leaderboard_data = {}
        self.all_averages = {}
        self.scrape_failures = {}

        # Local directory relative to the script location (same pattern as naukri_updater.py)
        self.user_data_dir = os.path.abspath("brave_profile")
        os.makedirs(self.user_data_dir, exist_ok=True)

    async def setup_browser(self, playwright):
        print(f"Using user data directory: {self.user_data_dir}")
        self.browser = await playwright.chromium.launch_persistent_context(
            self.user_data_dir,
            executable_path=self.browser_executable_path,
            headless=self.headless,
            args=["--no-sandbox", "--disable-dev-shm-usage"],
        )
        print("Successfully launched Brave browser with persistent context.")

        # Get the first page, or create a new one if none exist
        self.page = self.browser.pages[0] if self.browser.pages else await self.browser.new_page()

        # Close any additional tabs that may have opened
        for extra_page in self.browser.pages[1:]:
            await extra_page.close()

    # ------------------------------------------------------------------
    # Login helper
    # ------------------------------------------------------------------

    async def wait_for_login(self, timeout=120):
        # ... (keep your existing check for being logged in)

        print("\n" + "=" * 60)
        print(f"MANUAL LOGIN REQUIRED (Timeout: {timeout} seconds)")
        print("Please log in to LinkedIn. The script will continue automatically.")
        print("=" * 60 + "\n")

        start_time = asyncio.get_event_loop().time()
        while asyncio.get_event_loop().time() - start_time < timeout:
            try:
                # Check if browser is still connected
                if not self.browser:
                    print("Browser context closed. Exiting.")
                    return False
                
                # Check for successful login
                url = self.page.url
                if "feed" in url or await self.page.query_selector("#global-nav"):
                    print("Login detected!")
                    await asyncio.sleep(1)
                    return True
            except Exception as e:
                # Instead of sys.exit(), print the error and wait
                print(f"Error checking login status: {e}")
            
            await asyncio.sleep(2) # Increased sleep to keep CPU usage low

        print("\nLogin timeout reached.")
        return False

    # ------------------------------------------------------------------
    # Page load helper
    # ------------------------------------------------------------------

    async def _load_page(self, url, wait_seconds=10):
        await self.page.goto(url, wait_until="domcontentloaded", timeout=wait_seconds * 1000)
        await asyncio.sleep(0.5)  # small settle buffer

    # ------------------------------------------------------------------
    # Extract user score + average from results page
    #
    # HTML structure (as of May 2026):
    #
    #   <div class="pr-golden-chiclet ...">
    #     <div class="pr-golden-chiclet__text">0:11</div>          <- user's score
    #     <div class="pr-golden-chiclet__subtext ...">
    #       with 0 redraws!
    #     </div>
    #     ...
    #     <div class="pr-golden-chiclet__subtext ...">
    #       Today's avg: 0:20                                       <- average
    #     </div>
    #   </div>
    # ------------------------------------------------------------------

    async def extract_results_page(self, game_key):
        """Returns (user_score, average) strings, either may be None."""
        user_score = None
        average = None

        try:
            await self.page.wait_for_selector(
                ".pr-connections-leaderboard-player__container, .pr-golden-chiclet",
                timeout=15000,
            )
        except PlaywrightTimeoutError:
            print("    Warning: Results elements not found on results page.")
            return None, None

        # Try to find "You" in the leaderboard preview on the results page
        try:
            containers = await self.page.query_selector_all(
                ".pr-connections-leaderboard-player__container"
            )
            for c in containers:
                try:
                    name_el = await c.query_selector(
                        ".pr-connections-leaderboard-player__name"
                    )
                    if name_el and (await name_el.inner_text()).strip() == "You":
                        score_el = await c.query_selector(
                            ".pr-connections-leaderboard-player__score"
                        )
                        if score_el:
                            raw = clean_score((await score_el.inner_text()).strip())
                            if raw and raw != "--":
                                user_score = raw
                                print(f"    User score (results page): {user_score}")
                                break
                except Exception:
                    continue
        except Exception:
            pass

        if not user_score:
            try:
                raw = await self.page.evaluate("""() => {
                    const subtexts = document.querySelectorAll('.pr-golden-chiclet__subtext');
                    for (const st of subtexts) {
                        const txt = (st.innerText || '').toLowerCase();
                        if (!txt.includes('avg:')) continue;
                        let sib = st.previousElementSibling;
                        while (sib) {
                            if (sib.classList.contains('pr-golden-chiclet__text')) {
                                return (sib.innerText || '').trim();
                            }
                            sib = sib.previousElementSibling;
                        }
                    }
                    return null;
                }""")
                if raw:
                    raw = clean_score(raw)
                    m = re.search(r"solved in (\d+)", raw, re.IGNORECASE)
                    if m:
                        raw = m.group(1)
                    if raw and raw != "--":
                        user_score = raw
                        print(f"    User score (results page fallback): {user_score}")
            except Exception:
                pass

        try:
            subtexts = await self.page.query_selector_all(".pr-golden-chiclet__subtext")
            for st in subtexts:
                try:
                    text = (await st.inner_text() or "").strip()
                except Exception:
                    continue
                if "avg:" in text.lower():
                    idx = text.lower().find("avg:")
                    if idx != -1:
                        avg_raw = text[idx + 4:].strip()
                        avg_raw = clean_score(avg_raw)
                        if avg_raw and avg_raw != "--":
                            average = avg_raw
                            print(f"    Average (results page): {average}")
                    break
        except Exception:
            pass

        return user_score, average

    # ------------------------------------------------------------------
    # Extract leaderboard from connections page
    # ------------------------------------------------------------------

    async def extract_leaderboard_data(self, game_key):
        print(f"    Extracting leaderboard entries...")

        try:
            await self.page.wait_for_selector(
                ".pr-connections-leaderboard-player__container",
                timeout=20000,
            )
        except PlaywrightTimeoutError:
            print("    Warning: No leaderboard containers found.")
            return []

        script = """(gameKey) => {
            const isTimed   = ['mini-sudoku','zip','crossclimb','queens','tango','patches'].includes(gameKey);
            const isPinpoint = gameKey === 'pinpoint';

            function parseScore(raw) {
                if (!raw) return null;
                raw = raw.replace(/^['\u2019`"]+/, '').trim();
                if (!raw || raw === '--' || raw === '-') return null;
                if (/^\\d+:\\d{2}$/.test(raw)) return raw;
                if (/^\\d+$/.test(raw)) {
                    if (isPinpoint) return raw;
                    if (isTimed) return null;
                    return raw;
                }
                return null;
            }

            const contentSection = document.querySelector('.pr-connections-leaderboard__content');
            const containers = contentSection
                ? Array.from(contentSection.querySelectorAll('.pr-connections-leaderboard-player__container'))
                : Array.from(document.querySelectorAll('.pr-connections-leaderboard-player__container'));

            const players = [];
            containers.forEach(el => {
                const nameEl = el.querySelector(
                    '.pr-connections-leaderboard-player__content-column .text-body-medium-bold'
                );
                const name = nameEl ? nameEl.innerText.trim() : null;
                if (!name) return;

                const scoreEl = el.querySelector(
                    '.pr-connections-leaderboard-player__score.text-body-medium.ml1, ' +
                    '.pr-connections-leaderboard-player__score'
                );
                const score = scoreEl ? parseScore(scoreEl.innerText.trim()) : null;
                if (!score) return;

                players.push({ name, score });
            });

            return players;
        }"""

        try:
            raw_players = await self.page.evaluate(script, game_key)
        except Exception as e:
            print(f"    Warning: JS extraction failed: {e}")
            return []

        players = []
        for p in raw_players:
            if p and p.get("name") and p.get("score"):
                players.append(
                    {
                        "rank": len(players) + 1,
                        "name": p["name"].strip(),
                        "score": p["score"].strip(),
                    }
                )
                print(
                    f"      {players[-1]['rank']}. {players[-1]['name']} -- {players[-1]['score']}"
                )

        return players

    # ------------------------------------------------------------------
    # Scrape a single game
    # Returns True on success, False on failure
    # ------------------------------------------------------------------

    # ------------------------------------------------------------------
    # Helper to load URL with retries & page reload on failure
    # ------------------------------------------------------------------

    async def _goto_with_retry(self, page, url, game_name, max_retries=2, timeout=12000):
        for attempt in range(1, max_retries + 1):
            try:
                await page.goto(url, wait_until="domcontentloaded", timeout=timeout)
                await asyncio.sleep(0.5)
                return True
            except Exception as e:
                print(f"    [{game_name}] Navigation attempt {attempt}/{max_retries} failed for {url}: {e}")
                if attempt < max_retries:
                    await asyncio.sleep(1)
                    try:
                        print(f"    [{game_name}] Reloading page (attempt {attempt + 1})...")
                        await page.reload(wait_until="domcontentloaded", timeout=timeout)
                        await asyncio.sleep(0.5)
                        return True
                    except Exception as re_err:
                        print(f"    [{game_name}] Reload failed: {re_err}")
        return False

    # ------------------------------------------------------------------
    # Scrape a single game
    # Returns (game_name, success, players, avg, failures)
    # ------------------------------------------------------------------

    async def scrape_game(self, page, game_key, game_name):
        print(f"\n-- {game_name} --")
        failures = []
        avg_found = None
        players = []

        # Step 1: Results page -> user score + average
        results_url = RESULTS_URL.format(game=game_key)
        print(f"  [{game_name}] Loading results page: {results_url}")
        user_score_results = None
        try:
            await self._goto_with_retry(page, results_url, game_name)

            # Extract results page elements with retry/reload if missing selector
            try:
                await page.wait_for_selector(
                    ".pr-connections-leaderboard-player__container, .pr-golden-chiclet",
                    timeout=15000,
                )
            except PlaywrightTimeoutError:
                print(f"    [{game_name}] Elements not found on results page. Retrying page reload...")
                try:
                    await page.reload(wait_until="domcontentloaded", timeout=12000)
                    await page.wait_for_selector(
                        ".pr-connections-leaderboard-player__container, .pr-golden-chiclet",
                        timeout=12000,
                    )
                except Exception:
                    print(f"    [{game_name}] Warning: Results elements still not found after reload.")

            # Try to find "You" in the leaderboard preview on the results page
            try:
                containers = await page.query_selector_all(
                    ".pr-connections-leaderboard-player__container"
                )
                for c in containers:
                    try:
                        name_el = await c.query_selector(
                            ".pr-connections-leaderboard-player__name"
                        )
                        if name_el and (await name_el.inner_text()).strip() == "You":
                            score_el = await c.query_selector(
                                ".pr-connections-leaderboard-player__score"
                            )
                            if score_el:
                                raw = clean_score((await score_el.inner_text()).strip())
                                if raw and raw != "--":
                                    user_score_results = raw
                                    print(f"    [{game_name}] User score (results page): {user_score_results}")
                                    break
                    except Exception:
                        continue
            except Exception:
                pass

            if not user_score_results:
                try:
                    raw = await page.evaluate("""() => {
                        const subtexts = document.querySelectorAll('.pr-golden-chiclet__subtext');
                        for (const st of subtexts) {
                            const txt = (st.innerText || '').toLowerCase();
                            if (!txt.includes('avg:')) continue;
                            let sib = st.previousElementSibling;
                            while (sib) {
                                if (sib.classList.contains('pr-golden-chiclet__text')) {
                                    return (sib.innerText || '').trim();
                                }
                                sib = sib.previousElementSibling;
                            }
                        }
                        return null;
                    }""")
                    if raw:
                        raw = clean_score(raw)
                        m = re.search(r"solved in (\d+)", raw, re.IGNORECASE)
                        if m:
                            raw = m.group(1)
                        if raw and raw != "--":
                            user_score_results = raw
                            print(f"    [{game_name}] User score (results page fallback): {user_score_results}")
                except Exception:
                    pass

            try:
                subtexts = await page.query_selector_all(".pr-golden-chiclet__subtext")
                for st in subtexts:
                    try:
                        text = (await st.inner_text() or "").strip()
                    except Exception:
                        continue
                    if "avg:" in text.lower():
                        idx = text.lower().find("avg:")
                        if idx != -1:
                            avg_raw = text[idx + 4:].strip()
                            avg_raw = clean_score(avg_raw)
                            if avg_raw and avg_raw != "--":
                                avg_found = avg_raw
                                print(f"    [{game_name}] Average (results page): {avg_found}")
                        break
            except Exception:
                pass

            if avg_found:
                self.all_averages[game_name] = avg_found
            else:
                failures.append("average not found on results page")

            if not user_score_results:
                failures.append("personal score not found on results page")

        except Exception as e:
            print(f"  [{game_name}] Warning: Error on results page: {e}")
            failures.append(f"results page exception: {e}")

        # Step 2: Leaderboard page -> all ranked players
        lb_url = LEADERBOARD_URL.format(game=game_key)
        print(f"  [{game_name}] Loading leaderboard: {lb_url}")
        try:
            await self._goto_with_retry(page, lb_url, game_name)

            print(f"    [{game_name}] Extracting leaderboard entries...")
            try:
                await page.wait_for_selector(
                    ".pr-connections-leaderboard-player__container",
                    timeout=15000,
                )
            except PlaywrightTimeoutError:
                print(f"    [{game_name}] Leaderboard containers missing. Retrying page reload...")
                try:
                    await page.reload(wait_until="domcontentloaded", timeout=15000)
                    await page.wait_for_selector(
                        ".pr-connections-leaderboard-player__container",
                        timeout=15000,
                    )
                except Exception:
                    print(f"    [{game_name}] Warning: Leaderboard containers still missing after reload.")

            script = """(gameKey) => {
                const isTimed   = ['mini-sudoku','zip','crossclimb','queens','tango','patches'].includes(gameKey);
                const isPinpoint = gameKey === 'pinpoint';

                function parseScore(raw) {
                    if (!raw) return null;
                    raw = raw.replace(/^['\u2019`"]+/, '').trim();
                    if (!raw || raw === '--' || raw === '-') return null;
                    if (/^\\d+:\\d{2}$/.test(raw)) return raw;
                    if (/^\\d+$/.test(raw)) {
                        if (isPinpoint) return raw;
                        if (isTimed) return null;
                        return raw;
                    }
                    return null;
                }

                const contentSection = document.querySelector('.pr-connections-leaderboard__content');
                const containers = contentSection
                    ? Array.from(contentSection.querySelectorAll('.pr-connections-leaderboard-player__container'))
                    : Array.from(document.querySelectorAll('.pr-connections-leaderboard-player__container'));

                const players = [];
                containers.forEach(el => {
                    const nameEl = el.querySelector(
                        '.pr-connections-leaderboard-player__content-column .text-body-medium-bold'
                    );
                    const name = nameEl ? nameEl.innerText.trim() : null;
                    if (!name) return;

                    const scoreEl = el.querySelector(
                        '.pr-connections-leaderboard-player__score.text-body-medium.ml1, ' +
                        '.pr-connections-leaderboard-player__score'
                    );
                    const score = scoreEl ? parseScore(scoreEl.innerText.trim()) : null;
                    if (!score) return;

                    players.push({ name, score });
                });

                return players;
            }"""

            try:
                raw_players = await page.evaluate(script, game_key)
            except Exception as e:
                print(f"    [{game_name}] Warning: JS extraction failed: {e}")
                raw_players = []

            for p in raw_players:
                if p and p.get("name") and p.get("score"):
                    players.append(
                        {
                            "rank": len(players) + 1,
                            "name": p["name"].strip(),
                            "score": p["score"].strip(),
                        }
                    )
                    print(
                        f"      [{game_name}] {players[-1]['rank']}. {players[-1]['name']} -- {players[-1]['score']}"
                    )
        except Exception as e:
            print(f"  [{game_name}] Warning: Error on leaderboard page: {e}")
            players = []

        # Reconcile "You" between the two pages.
        # LinkedIn's leaderboard page often renders the "You" row twice (a
        # pinned/sticky copy plus the row at your actual rank), so `players`
        # can contain more than one "You" entry here. Collapse them to a
        # single entry -- this is what was causing rows to multiply on every
        # merge in export_to_excel.
        you_entries = [p for p in players if p["name"].lower() == "you"]
        if you_entries:
            you_entry = you_entries[0]
            if len(you_entries) > 1:
                print(
                    f"  [{game_name}] Note: Found {len(you_entries)} 'You' rows on the "
                    f"leaderboard page; collapsing to one."
                )
            if user_score_results and you_entry["score"] != user_score_results:
                print(
                    f"  [{game_name}] Note: Using leaderboard score ({you_entry['score']}) for 'You' "
                    f"(results page had {user_score_results})"
                )
            # Drop every "You" row, then re-add exactly one.
            players = [p for p in players if p["name"].lower() != "you"]
            players.insert(0, you_entry)
        elif user_score_results:
            print(f"  [{game_name}] Adding 'You' from results page: {user_score_results}")
            players.insert(0, {"rank": 0, "name": "You", "score": user_score_results})

        # Safety net: also dedupe by name generally (any other player whose
        # row got matched by more than one container), keeping the first
        # occurrence, since names are otherwise treated as the merge key.
        seen_names = set()
        deduped_players = []
        for p in players:
            key = p["name"].strip().lower()
            if key in seen_names:
                continue
            seen_names.add(key)
            deduped_players.append(p)
        players = deduped_players

        if failures:
            self.scrape_failures[game_name] = failures

        if players:
            self.all_leaderboard_data[game_name] = players
            print(f"  [{game_name}] OK: {len(players)} player(s) recorded.")
            return True
        else:
            print(f"  [{game_name}] FAIL: No data found.")
            return False

    # ------------------------------------------------------------------
    # Scrape all games -- pinpoint first on main tab, then parallel tabs
    # ------------------------------------------------------------------

    async def scrape_all_games(self):
        print("\n" + "=" * 60)
        print("Starting leaderboard scraping")
        print("=" * 60)

        game_items = list(GAMES.items())
        pinpoint_key, pinpoint_name = game_items[0]

        # 1. Scrape Pinpoint first on the single primary page
        success = await self.scrape_game(self.page, pinpoint_key, pinpoint_name)
        if not success:
            print(
                "\nFAIL: Pinpoint scrape failed. This usually means LinkedIn's page "
                "structure has changed or you are not logged in."
            )
            print("  Exiting early -- no data written.")
            self._print_failure_summary()
            return False

        # 2. Scrape all remaining games concurrently in separate browser tabs
        remaining_games = game_items[1:]
        print(f"\nPinpoint successful. Scraping remaining {len(remaining_games)} games in parallel tabs...")

        async def worker(game_key, game_name):
            tab = await self.browser.new_page()
            try:
                return await self.scrape_game(tab, game_key, game_name)
            finally:
                await tab.close()

        tasks = [worker(k, v) for k, v in remaining_games]
        await asyncio.gather(*tasks)

        self._print_failure_summary()
        return True

    # ------------------------------------------------------------------
    # Print a summary of any scrape failures
    # ------------------------------------------------------------------

    def _print_failure_summary(self):
        if not self.scrape_failures:
            return
        print("\n" + "=" * 60)
        print("SCRAPE FAILURE SUMMARY (elements may have changed)")
        print("=" * 60)
        for game_name, issues in self.scrape_failures.items():
            print(f"  {game_name}:")
            for issue in issues:
                print(f"    - {issue}")
        print("=" * 60)

    # ------------------------------------------------------------------
    # Export to Excel (monthly rolling file)
    # ------------------------------------------------------------------

    def export_to_excel(self):
        # --- Added setup for the folder ---
        folder_name = "sheets"
        if not os.path.exists(folder_name):
            os.makedirs(folder_name)
        # ----------------------------------

        now = datetime.datetime.now()
        if now.hour < 13 or (now.hour == 13 and now.minute < 30):
            report_date = now - datetime.timedelta(days=1)
        else:
            report_date = now

        date_str = report_date.strftime("%Y-%m-%d")
        month_str = report_date.strftime("%b-%Y")
        
        # Construct full path
        file_name = f"linkedin_leaderboards_{month_str}.xlsx"
        file_path = os.path.join(folder_name, file_name)

        print(f"\nUpdating {file_path} for {date_str}...")

        all_sheets = {}
        if os.path.exists(file_path):
            try:
                with pd.ExcelFile(file_path) as xls:
                    for s_name in xls.sheet_names:
                        all_sheets[s_name] = pd.read_excel(xls, sheet_name=s_name)
            except Exception as e:
                print(f"  Warning: could not read {file_path}: {e}. Starting fresh.")

        for game_name, players in self.all_leaderboard_data.items():
            if not players:
                continue

            new_records = []
            if game_name in self.all_averages:
                new_records.append(
                    {"Player Name": "Average", date_str: self.all_averages[game_name]}
                )

            for p in players:
                new_records.append(
                    {"Player Name": p["name"].strip(), date_str: p["score"]}
                )

            new_df = pd.DataFrame(new_records)
            # Safety net: pd.merge on a key that isn't unique on both sides
            # does a Cartesian match (1 row x 2 rows -> 2 rows, 2x2 -> 4, ...),
            # which is how duplicate "Player Name" rows compounded over time.
            # Dedupe today's records before they ever reach the merge.
            new_df = dedupe_by_player_name(new_df)

            if game_name in all_sheets:
                df = all_sheets[game_name].copy()
                df["Player Name"] = df["Player Name"].astype(str).str.strip()
                # Also dedupe the existing sheet on load, so any duplicates
                # already sitting in the file (from before this fix) can't
                # keep multiplying on subsequent runs.
                df = dedupe_by_player_name(df)
                if date_str in df.columns:
                    df = df.drop(columns=[date_str])
                updated_df = pd.merge(df, new_df, on="Player Name", how="outer")
            else:
                updated_df = new_df

            all_sheets[game_name] = updated_df

        if not all_sheets:
            print("No data to export.")
            return

        # Use the full file_path here
        with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
            for g_name in sorted(all_sheets.keys()):
                df = all_sheets[g_name].copy()

                fixed_cols = ["Player Name"]
                date_cols = sorted([c for c in df.columns if c not in fixed_cols])
                df = df[fixed_cols + date_cols]

                df["_is_avg"] = df["Player Name"].str.strip().str.lower() == "average"
                df["_is_you"] = df["Player Name"].str.strip().str.lower() == "you"
                df = (
                    df.sort_values(
                        ["_is_avg", "_is_you", "Player Name"],
                        ascending=[False, False, True],
                    )
                    .drop(columns=["_is_avg", "_is_you"])
                    .reset_index(drop=True)
                )

                df.to_excel(writer, sheet_name=g_name, index=False)

                ws = writer.sheets[g_name]
                hdr_fill = PatternFill(
                    start_color="0066CC", end_color="0066CC", fill_type="solid"
                )
                hdr_font = Font(bold=True, color="FFFFFF")
                for cell in ws[1]:
                    cell.fill = hdr_fill
                    cell.font = hdr_font
                    cell.alignment = Alignment(horizontal="center")

                ws.column_dimensions["A"].width = 35
                for i in range(2, ws.max_column + 1):
                    ws.column_dimensions[get_column_letter(i)].width = 15

        print(f"OK: Excel file updated: {file_path}")

    # ------------------------------------------------------------------
    # Entry point
    # ------------------------------------------------------------------

    async def run(self):
        async with async_playwright() as p:
            try:
                await self.setup_browser(p)
                
                # Navigate to LinkedIn
                print("Navigating to LinkedIn...")
                await self.page.goto("https://www.linkedin.com", timeout=60000)

                # Wait for login with a very generous timeout
                if not await self.wait_for_login(timeout=300): # 5 minutes
                    print("\nLogin failed or timed out.")
                    return # Exit the function, but keep the browser open for inspection

                # Proceed to scrape
                print("Scraping started...")
                ok = await self.scrape_all_games()
                if ok:
                    self.export_to_excel()
                    
            except Exception as e:
                print(f"\nCaught Exception: {e}")
                # We don't close the browser here so you can see what happened
            # No finally block here while debugging


# ---------------------------------------------------------------------------


def main():
    parser = argparse.ArgumentParser(description="LinkedIn Games Leaderboard Scraper")
    parser.add_argument(
        "--browser-path", type=str, help="Path to Brave/Chromium binary (default: /usr/bin/brave)"
    )
    parser.add_argument(
        "--headless", action="store_true", help="Run browser in headless mode"
    )
    args = parser.parse_args()

    scraper = LinkedInLeaderboardScraper(
        browser_executable_path=args.browser_path,
        headless=args.headless,
    )
    asyncio.run(scraper.run())


if __name__ == "__main__":
    main()
