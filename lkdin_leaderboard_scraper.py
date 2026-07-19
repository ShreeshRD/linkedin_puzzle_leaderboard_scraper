"""
LinkedIn Games Leaderboard Scraper
Automates browser to fetch leaderboard data from LinkedIn games and exports to Excel.

Fix changelog:
- Pinpoint scraped first; script exits early if it fails (confirms login + page structure is working)
- Results page scraped first for user score + average
- Average extracted from .pr-golden-chiclet__subtext containing "Today's avg:"
- User score extracted from .pr-golden-chiclet__text (the big number/time shown on results page)
- Leaderboard scrape targets .pr-connections-leaderboard__sticky-section for ranked players
  AND falls back to all .pr-connections-leaderboard-player__container elements
- Fixed "You" detection: name is in .pr-connections-leaderboard-player__content-column .text-body-medium-bold
  (the old .pr-connections-leaderboard-player__text-wrapper wrapper no longer exists in the DOM)
"""

import time
import re
import argparse
import os
import sys
import datetime
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# Game order: pinpoint MUST be first so we can bail early on failure
# ---------------------------------------------------------------------------
GAMES = {
    'pinpoint':    'Pinpoint',
    'wend':        'Wend',
    'mini-sudoku': 'Mini Sudoku',
    'zip':         'Zip',
    'crossclimb':  'Crossclimb',
    'queens':      'Queens',
    'tango':       'Tango',
    'patches':     'Patches',
}

# TIMED_GAMES = {'mini-sudoku', 'zip', 'crossclimb', 'queens', 'tango', 'patches', 'wend'}

RESULTS_URL    = "https://www.linkedin.com/games/{game}/results/"
LEADERBOARD_URL = "https://www.linkedin.com/games/{game}/results/leaderboard/connections/"


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def clean_score(text: str) -> str:
    """Strip leading apostrophes/backticks that LinkedIn sometimes injects."""
    if not text:
        return ""
    return re.sub(r"^['\"`\u2019]+", "", text.strip()).strip()


def parse_time_seconds(score: str) -> float:
    """Convert M:SS to float seconds for sorting. Returns large number on failure."""
    m = re.match(r'^(\d+):(\d{2})$', score.strip())
    if m:
        return int(m.group(1)) * 60 + int(m.group(2))
    m2 = re.match(r'^(\d+)$', score.strip())
    if m2:
        return float(m2.group(1))
    return 1e9


# ---------------------------------------------------------------------------
# Main scraper class
# ---------------------------------------------------------------------------

class LinkedInLeaderboardScraper:
    def __init__(self, browser_executable_path=None, headless=False):
        self.browser_executable_path = browser_executable_path
        self.headless = headless
        self.driver = None
        self.all_leaderboard_data = {}
        self.all_averages = {}
        self.scrape_failures = {}   # game_name -> list of failure strings
        self.user_data_dir = os.path.abspath(".chrome_profile")
        os.makedirs(self.user_data_dir, exist_ok=True)

    # ------------------------------------------------------------------
    # Browser setup
    # ------------------------------------------------------------------

    def setup_driver(self):
        options = webdriver.ChromeOptions()
        options.page_load_strategy = 'eager'
        if self.browser_executable_path:
            options.binary_location = self.browser_executable_path
        if self.headless:
            options.add_argument('--headless')
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-dev-shm-usage')
        options.add_argument('--remote-debugging-port=9222')
        options.add_argument(f'--user-data-dir={self.user_data_dir}')
        options.add_argument('--profile-directory=Default')
        options.add_argument('--start-maximized')

        print("Installing/Updating ChromeDriver...")
        driver_path = ChromeDriverManager().install()
        service = ChromeService(driver_path)
        self.driver = webdriver.Chrome(service=service, options=options)
        print("Browser initialized successfully")

    # ------------------------------------------------------------------
    # Login helper
    # ------------------------------------------------------------------

    def wait_for_login(self, timeout=120):
        try:
            if "feed" in self.driver.current_url or self.driver.find_elements(By.ID, "global-nav"):
                print("Already logged in.")
                return True
        except Exception:
            pass

        print("\n" + "=" * 60)
        print("MANUAL LOGIN REQUIRED (Timeout: 120 seconds)")
        print("=" * 60)
        print("Please log in to LinkedIn in the browser window.")
        print("The script will automatically continue once logged in.")
        print("=" * 60 + "\n")

        start = time.time()
        while time.time() - start < timeout:
            try:
                if not self.driver.window_handles:
                    print("Browser closed. Exiting...")
                    sys.exit(0)
                if "feed" in self.driver.current_url or self.driver.find_elements(By.ID, 'global-nav'):
                    print("Login detected!")
                    time.sleep(0.5)
                    return True
            except Exception:
                print("\nBrowser connection lost. Exiting...")
                sys.exit(0)
            time.sleep(0.5)

        print("\nLogin timeout reached.")
        return False

    # ------------------------------------------------------------------
    # Page load helper
    # ------------------------------------------------------------------

    def _load_page(self, url, wait_seconds=10):
        self.driver.get(url)
        WebDriverWait(self.driver, wait_seconds).until(
            lambda d: d.execute_script('return document.readyState') == 'complete'
        )
        time.sleep(0.5)  # small settle buffer

    # ------------------------------------------------------------------
    # Extract user score + average from results page
    #
    # HTML structure (as of May 2026):
    #
    #   <div class="pr-golden-chiclet ...">
    #     <div class="pr-golden-chiclet__text">0:11</div>          ← user's score
    #     <div class="pr-golden-chiclet__subtext ...">
    #       with 0 redraws!
    #     </div>
    #     ...
    #     <div class="pr-golden-chiclet__subtext ...">
    #       Today's avg: 0:20                                       ← average
    #     </div>
    #   </div>
    # ------------------------------------------------------------------

    def extract_results_page(self, game_key):
        """Returns (user_score, average) strings, either may be None."""
        user_score = None
        average = None

        try:
            # Wait for either the new leaderboard container or the old chiclet
            WebDriverWait(self.driver, 15).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, '.pr-connections-leaderboard-player__container, .pr-golden-chiclet'))
            )
        except Exception:
            print("    ⚠ Results elements not found on results page.")
            return None, None

        # Try to find "You" in the leaderboard preview on the results page
        try:
            containers = self.driver.find_elements(By.CSS_SELECTOR, '.pr-connections-leaderboard-player__container')
            for c in containers:
                try:
                    name_el = c.find_element(By.CSS_SELECTOR, '.pr-connections-leaderboard-player__name')
                    if name_el.text.strip() == 'You':
                        score_el = c.find_element(By.CSS_SELECTOR, '.pr-connections-leaderboard-player__score')
                        raw = clean_score(score_el.text.strip())
                        if raw and raw != '--':
                            user_score = raw
                            print(f"    User score (results page): {user_score}")
                            break
                except Exception:
                    continue
        except Exception:
            pass

        if not user_score:
            try:
                # Fallback: chiclet carousel — find the __text slide whose immediately
                # following __subtext sibling contains "avg:".  That sibling pair is the
                # score slide.  All slides share the same class, so we must use JS to
                # walk siblings rather than grabbing the first element.
                raw = self.driver.execute_script("""
                    const subtexts = document.querySelectorAll('.pr-golden-chiclet__subtext');
                    for (const st of subtexts) {
                        const txt = (st.innerText || '').toLowerCase();
                        if (!txt.includes('avg:')) continue;
                        // Walk backwards through siblings to find the preceding __text
                        let sib = st.previousElementSibling;
                        while (sib) {
                            if (sib.classList.contains('pr-golden-chiclet__text')) {
                                return (sib.innerText || '').trim();
                            }
                            sib = sib.previousElementSibling;
                        }
                    }
                    return null;
                """)
                if raw:
                    raw = clean_score(raw)
                    # "Solved in N" → extract the trailing integer (Pinpoint)
                    m = re.search(r'solved in (\d+)', raw, re.IGNORECASE)
                    if m:
                        raw = m.group(1)
                    if raw and raw != '--':
                        user_score = raw
                        print(f"    User score (results page fallback): {user_score}")
            except Exception:
                pass

        try:
            # Average: the subtext element that contains "avg:"
            # NOTE: LinkedIn uses a carousel that positions slides off-screen via CSS
            # transform. Selenium's .text returns '' for off-screen elements, so we
            # use get_attribute('innerText') which reads DOM text regardless of position.
            subtexts = self.driver.find_elements(By.CSS_SELECTOR, '.pr-golden-chiclet__subtext')
            for st in subtexts:
                try:
                    text = (st.get_attribute('innerText') or '').strip()
                except Exception:
                    continue
                if 'avg:' in text.lower():
                    idx = text.lower().find('avg:')
                    if idx != -1:
                        avg_raw = text[idx + 4:].strip()
                        avg_raw = clean_score(avg_raw)
                        if avg_raw and avg_raw != '--':
                            average = avg_raw
                            print(f"    Average (results page): {average}")
                    break
        except Exception:
            pass

        return user_score, average

    # ------------------------------------------------------------------
    # Extract leaderboard from connections page
    #
    # Two groups exist in the DOM:
    #   1. .pr-connections-leaderboard__sticky-section  → top ranked players
    #      (shown at top, may contain "You" and your friends)
    #   2. All remaining .pr-connections-leaderboard-player__container elements
    #      (players who have played but ranked lower, OR "Nudge to play" section)
    #
    # We want group 1 only (players with scores).  The "Nudge to play" section
    # contains containers WITHOUT a score element, so we can filter them out.
    #
    # Name is in: .pr-connections-leaderboard-player__content-column .text-body-medium-bold
    # Score is in: .pr-connections-leaderboard-player__score.text-body-medium.ml1
    # ------------------------------------------------------------------

    def extract_leaderboard_data(self, game_key):
        print(f"    Extracting leaderboard entries...")

        # Wait for any container to appear
        try:
            WebDriverWait(self.driver, 20).until(
                EC.presence_of_element_located(
                    (By.CSS_SELECTOR, '.pr-connections-leaderboard-player__container')
                )
            )
        except Exception:
            print("    ⚠ No leaderboard containers found.")
            return []

        # Use JS for bulk extraction — fast and avoids stale-element issues
        script = r"""
        const gameKey   = arguments[0];
        const isTimed   = ['mini-sudoku','zip','crossclimb','queens','tango','patches'].includes(gameKey);
        const isPinpoint = gameKey === 'pinpoint';

        // Helper: parse a raw score string into a canonical form, or null if not valid
        function parseScore(raw) {
            if (!raw) return null;
            // strip leading apostrophes LinkedIn sometimes injects
            raw = raw.replace(/^['\u2019`"]+/, '').trim();
            if (!raw || raw === '--' || raw === '-') return null;

            // Time format M:SS
            if (/^\d+:\d{2}$/.test(raw)) return raw;

            // Pure integer (used by Pinpoint, 1–5)
            if (/^\d+$/.test(raw)) {
                if (isPinpoint) return raw;
                // For timed games, a bare integer is likely a streak counter — reject it
                if (isTimed) return null;
                return raw;
            }
            return null;
        }

        // Prefer containers inside the ranked content section
        // but fall back to ALL containers if content section is not found.
        const contentSection = document.querySelector('.pr-connections-leaderboard__content');
        const containers = contentSection
            ? Array.from(contentSection.querySelectorAll('.pr-connections-leaderboard-player__container'))
            : Array.from(document.querySelectorAll('.pr-connections-leaderboard-player__container'));

        const players = [];
        containers.forEach(el => {
            // Name: directly in .pr-connections-leaderboard-player__content-column > .text-body-medium-bold
            // (the old __text-wrapper wrapper no longer exists)
            const nameEl = el.querySelector(
                '.pr-connections-leaderboard-player__content-column .text-body-medium-bold'
            );
            const name = nameEl ? nameEl.innerText.trim() : null;
            if (!name) return;

            // Score: the element with class pr-connections-leaderboard-player__score AND ml1
            const scoreEl = el.querySelector(
                '.pr-connections-leaderboard-player__score.text-body-medium.ml1, ' +
                '.pr-connections-leaderboard-player__score'
            );
            const score = scoreEl ? parseScore(scoreEl.innerText.trim()) : null;
            if (!score) return;  // skip nudge-section entries (no score element)

            players.push({ name, score });
        });

        return players;
        """

        try:
            raw_players = self.driver.execute_script(script, game_key)
        except Exception as e:
            print(f"    ⚠ JS extraction failed: {e}")
            return []

        players = []
        for p in raw_players:
            if p and p.get('name') and p.get('score'):
                players.append({
                    'rank': len(players) + 1,
                    'name': p['name'].strip(),
                    'score': p['score'].strip(),
                })
                print(f"      {players[-1]['rank']}. {players[-1]['name']} — {players[-1]['score']}")

        return players

    # ------------------------------------------------------------------
    # Scrape a single game
    # Returns True on success, False on failure
    # ------------------------------------------------------------------

    def scrape_game(self, game_key, game_name):
        print(f"\n── {game_name} ──")
        failures = []

        # Step 1: Results page → user score + average
        results_url = RESULTS_URL.format(game=game_key)
        print(f"  Loading results page: {results_url}")
        user_score_results = None
        try:
            self._load_page(results_url)
            user_score_results, avg = self.extract_results_page(game_key)
            if avg:
                self.all_averages[game_name] = avg
            else:
                failures.append("average not found on results page")
            if not user_score_results:
                failures.append("personal score not found on results page")
        except Exception as e:
            print(f"  ⚠ Error on results page: {e}")
            failures.append(f"results page exception: {e}")

        # Step 2: Leaderboard page → all ranked players
        lb_url = LEADERBOARD_URL.format(game=game_key)
        print(f"  Loading leaderboard: {lb_url}")
        try:
            self._load_page(lb_url)
            players = self.extract_leaderboard_data(game_key)
        except Exception as e:
            print(f"  ⚠ Error on leaderboard page: {e}")
            players = []

        # Reconcile "You" between the two pages
        you_entry = next((p for p in players if p['name'].lower() == 'you'), None)
        if you_entry:
            if user_score_results and you_entry['score'] != user_score_results:
                print(f"  Note: Using leaderboard score ({you_entry['score']}) for 'You' "
                      f"(results page had {user_score_results})")
        elif user_score_results:
            print(f"  Adding 'You' from results page: {user_score_results}")
            players.insert(0, {'rank': 0, 'name': 'You', 'score': user_score_results})

        # Record any failures for this game
        if failures:
            self.scrape_failures[game_name] = failures

        if players:
            self.all_leaderboard_data[game_name] = players
            print(f"  ✓ {len(players)} player(s) recorded for {game_name}.")
            return True
        else:
            print(f"  ✗ No data found for {game_name}.")
            return False

    # ------------------------------------------------------------------
    # Scrape all games — pinpoint first, exit early on failure
    # ------------------------------------------------------------------

    def scrape_all_games(self):
        print("\n" + "=" * 60)
        print("Starting leaderboard scraping")
        print("=" * 60)

        game_items = list(GAMES.items())  # pinpoint is first by construction

        for i, (game_key, game_name) in enumerate(game_items):
            success = self.scrape_game(game_key, game_name)

            # If Pinpoint (index 0) fails, bail out entirely
            if i == 0 and not success:
                print("\n✗ Pinpoint scrape failed. This usually means LinkedIn's page "
                      "structure has changed or you are not logged in.")
                print("  Exiting early — no data written.")
                self._print_failure_summary()
                return False

            time.sleep(1.5)  # polite delay between games

        self._print_failure_summary()
        return True

    # ------------------------------------------------------------------
    # Print a summary of any scrape failures (avg / personal score)
    # ------------------------------------------------------------------

    def _print_failure_summary(self):
        if not self.scrape_failures:
            return
        print("\n" + "=" * 60)
        print("⚠  SCRAPE FAILURE SUMMARY (elements may have changed)")
        print("=" * 60)
        for game_name, issues in self.scrape_failures.items():
            print(f"  {game_name}:")
            for issue in issues:
                print(f"    • {issue}")
        print("=" * 60)

    # ------------------------------------------------------------------
    # Export to Excel (monthly rolling file)
    # ------------------------------------------------------------------

    def export_to_excel(self):
        # Use yesterday's date if run before 1:30 PM
        now = datetime.datetime.now()
        if now.hour < 13 or (now.hour == 13 and now.minute < 30):
            report_date = now - datetime.timedelta(days=1)
        else:
            report_date = now

        date_str   = report_date.strftime('%Y-%m-%d')
        month_str  = report_date.strftime('%b-%Y')
        monthly_file = f"linkedin_leaderboards_{month_str}.xlsx"

        print(f"\nUpdating {monthly_file} for {date_str}...")

        # Load existing sheets
        all_sheets = {}
        if os.path.exists(monthly_file):
            try:
                with pd.ExcelFile(monthly_file) as xls:
                    for s_name in xls.sheet_names:
                        all_sheets[s_name] = pd.read_excel(xls, sheet_name=s_name)
            except Exception as e:
                print(f"  Warning: could not read {monthly_file}: {e}. Starting fresh.")

        # Merge today's data into each game's sheet
        for game_name, players in self.all_leaderboard_data.items():
            if not players:
                continue

            new_records = []

            if game_name in self.all_averages:
                new_records.append({'Player Name': 'Average', date_str: self.all_averages[game_name]})

            for p in players:
                new_records.append({'Player Name': p['name'].strip(), date_str: p['score']})

            new_df = pd.DataFrame(new_records)

            if game_name in all_sheets:
                df = all_sheets[game_name].copy()
                df['Player Name'] = df['Player Name'].astype(str).str.strip()
                if date_str in df.columns:
                    df = df.drop(columns=[date_str])
                updated_df = pd.merge(df, new_df, on='Player Name', how='outer')
            else:
                updated_df = new_df

            all_sheets[game_name] = updated_df

        if not all_sheets:
            print("No data to export.")
            return

        with pd.ExcelWriter(monthly_file, engine='openpyxl') as writer:
            for g_name in sorted(all_sheets.keys()):
                df = all_sheets[g_name].copy()

                # Column order: Player Name, then date columns sorted ascending
                fixed_cols = ['Player Name']
                date_cols  = sorted([c for c in df.columns if c not in fixed_cols])
                df = df[fixed_cols + date_cols]

                # Row order: Average → You → others alphabetically
                df['_is_avg'] = df['Player Name'].str.strip().str.lower() == 'average'
                df['_is_you'] = df['Player Name'].str.strip().str.lower() == 'you'
                df = (df.sort_values(['_is_avg', '_is_you', 'Player Name'],
                                     ascending=[False, False, True])
                        .drop(columns=['_is_avg', '_is_you'])
                        .reset_index(drop=True))

                df.to_excel(writer, sheet_name=g_name, index=False)

                # Styling
                ws = writer.sheets[g_name]
                hdr_fill = PatternFill(start_color="0066CC", end_color="0066CC", fill_type="solid")
                hdr_font = Font(bold=True, color="FFFFFF")
                for cell in ws[1]:
                    cell.fill = hdr_fill
                    cell.font = hdr_font
                    cell.alignment = Alignment(horizontal='center')

                ws.column_dimensions['A'].width = 35
                for i in range(2, ws.max_column + 1):
                    ws.column_dimensions[get_column_letter(i)].width = 15

        print(f"✓ Excel file updated: {monthly_file}")

    # ------------------------------------------------------------------
    # Entry point
    # ------------------------------------------------------------------

    def run(self):
        try:
            self.setup_driver()
            self.driver.get("https://www.linkedin.com")
            if not self.wait_for_login():
                print("\nLogin failed or timed out. Exiting.")
                return
            ok = self.scrape_all_games()
            if ok:
                self.export_to_excel()
        except KeyboardInterrupt:
            print("\nInterrupted by user.")
        except Exception as e:
            import traceback
            print(f"\nUnexpected error: {e}")
            traceback.print_exc()
        finally:
            if self.driver:
                print("\nClosing browser...")
                self.driver.quit()


# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(description="LinkedIn Games Leaderboard Scraper")
    parser.add_argument('--browser-path', type=str, help="Path to Chrome/Chromium binary")
    args = parser.parse_args()

    scraper = LinkedInLeaderboardScraper(browser_executable_path=args.browser_path)
    scraper.run()


if __name__ == "__main__":
    main()