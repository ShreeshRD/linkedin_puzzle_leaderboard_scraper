"""
LinkedIn Games Leaderboard Scraper
Automates browser to fetch leaderboard data from LinkedIn games and exports to Excel
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

# Game configurations
GAMES = {
    'mini-sudoku': 'Mini Sudoku',
    'zip': 'Zip',
    'pinpoint': 'Pinpoint',
    'crossclimb': 'Crossclimb',
    'queens': 'Queens',
    'tango': 'Tango'
}

BASE_URL = "https://www.linkedin.com/games/{game}/results/leaderboard/connections/"


class LinkedInLeaderboardScraper:
    def __init__(self, browser_executable_path=None, headless=False):
        self.browser_executable_path = browser_executable_path
        self.headless = headless
        self.driver = None
        self.all_leaderboard_data = {}
        self.all_averages = {}
        # Persistent profile directory
        self.user_data_dir = os.path.abspath(".chrome_profile")
        
        if not os.path.exists(self.user_data_dir):
            os.makedirs(self.user_data_dir)
        
    def setup_driver(self):
        """Setup Chrome WebDriver with custom options and Service"""
        options = webdriver.ChromeOptions()
        options.page_load_strategy = 'eager'
        
        if self.browser_executable_path:
            options.binary_location = self.browser_executable_path
            
        if self.headless:
            options.add_argument('--headless')
            
        # Minimal flags for stability across WSL
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-dev-shm-usage')
        options.add_argument('--remote-debugging-port=9222')
        
        # Persistent profile
        options.add_argument(f'--user-data-dir={self.user_data_dir}')
        options.add_argument('--profile-directory=Default')
        
        # Keep browser open and logged in
        options.add_argument('--start-maximized')
        
        # Use Service for better driver management
        print("Installing/Updating ChromeDriver...")
        driver_path = ChromeDriverManager().install()
        service = ChromeService(driver_path)
        
        self.driver = webdriver.Chrome(service=service, options=options)
        print("Browser initialized successfully")
        
    def wait_for_login(self, timeout=120):
        # Quick check if already logged in to avoid unnecessary message
        try:
            # We check for the presence of the navigation bar or the feed URL
            if "feed" in self.driver.current_url or self.driver.find_elements(By.ID, "global-nav"):
                print("Login detected!")
                return True
        except Exception:
            print("\n" + "="*60)
            print("MANUAL LOGIN REQUIRED (Timeout: 120 seconds)")
            print("="*60)
            print("Please log in to LinkedIn in the browser window.")
            print("The script will automatically continue once you are logged in.")
            print("Or you can press ENTER here if you're already logged in.")
            print("If you close the browser, the script will exit gracefully.")
            print("="*60 + "\n")
            return None

        
        start_time = time.time()
        while time.time() - start_time < timeout:
            try:
                # Immediate check for browser closure
                if not self.driver.window_handles:
                    print("Browser closed by user. Exiting...")
                    sys.exit(0)
                
                # Check for logged-in indicators (URL or nav bar)
                # We use find_elements to avoid throwing exceptions if not found
                if "feed" in self.driver.current_url or self.driver.find_elements(By.ID, 'global-nav'):
                    print("Login detected!")
                    # Small buffer to ensure session cookies are fully set
                    time.sleep(0.5)
                    return True
                
            except Exception:
                print("\nBrowser connection lost. Exiting...")
                sys.exit(0)
                
            time.sleep(0.5) # Reduced from 2.0s for much faster detection
            
        print("\nLogin timeout reached.")
        return False
        
    def extract_average(self):
        """Extracts the average score from the results page subtext"""
        try:
            # Look for subtext elements which contain "avg:"
            subtexts = self.driver.find_elements(By.CSS_SELECTOR, '.pr-golden-chiclet__subtext')
            for st in subtexts:
                text = st.text.strip()
                # Pattern: "Today’s avg: 0:20" or "Today’s avg: 3"
                if 'avg:' in text.lower():
                    parts = text.split(':', 1)
                    if len(parts) > 1:
                        avg_val = parts[1].strip()
                        # Clean up leading apostrophes which LinkedIn often adds to scores (especially Pinpoint)
                        avg_val = re.sub(r"^['\"`]+", "", avg_val).strip()
                        # Simple validation: should look like a score or time
                        if avg_val and avg_val != "--":
                            return avg_val
        except:
            pass
        return None

    def extract_user_result(self, game_key):
        """Extracts the user's own score from the page"""
        is_timed = game_key in ['mini-sudoku', 'zip', 'crossclimb', 'queens', 'tango', 'patches']
        try:
            # 1. Try the new container structure (matches the HTML provided by user)
            containers = self.driver.find_elements(By.CSS_SELECTOR, '.pr-connections-leaderboard-player__container')
            for container in containers:
                name_elems = container.find_elements(By.CSS_SELECTOR, '.pr-connections-leaderboard-player__text-wrapper .text-body-medium-bold')
                if name_elems and name_elems[0].text.strip().lower() == 'you':
                    score_elems = container.find_elements(By.CSS_SELECTOR, '.pr-connections-leaderboard-player__score')
                    if score_elems:
                        score_text = score_elems[0].text.strip()
                        # Clean up leading apostrophes
                        score_text = re.sub(r"^['\"`]+", "", score_text).strip()
                        
                        if is_timed:
                            match = re.search(r'\d+:\d{2}', score_text)
                            if match: return match.group(0)
                        elif game_key == 'pinpoint':
                            match = re.search(r'\d+', score_text)
                            if match: return match.group(0)
                        elif score_text and score_text != "--":
                            return score_text

            # 2. Fallback to the big chiclet (results page style)
            # This handles cases like <div class="pr-golden-chiclet__text mt2">Solved in 3</div>
            elems = self.driver.find_elements(By.CSS_SELECTOR, '.pr-golden-chiclet__text')
            for el in elems:
                text = el.text.strip()
                text = re.sub(r"^['\"`]+", "", text).strip()
                
                if is_timed:
                    match = re.search(r'\d+:\d{2}', text)
                    if match: return match.group(0)
                elif game_key == 'pinpoint':
                    # Specifically look for a number, possibly in "Solved in 3"
                    match = re.search(r'\d+', text)
                    if match: return match.group(0)
                elif text and text != "--" and not any(word in text.lower() for word in ["fire", "streak"]):
                    return text
        except:
            pass
        return None

    def extract_leaderboard_data(self, game_key, game_name):
        print(f"  Extracting leaderboard for {game_name}...")
        try:
            # Wait for any element from the leaderboard to appear
            wait = WebDriverWait(self.driver, 20)
            wait.until(EC.presence_of_element_located(
                (By.CSS_SELECTOR, '.pr-connections-leaderboard-player__container')
            ))
        except Exception:
            print(f"    No leaderboard entries found for {game_name} or timed out.")
            return []
        
        # Bulk extract data using JavaScript to avoid per-row Selenium overhead
        # We pass game_key to help the script decide between timed and numeric scores
        script = """
        const gameKey = arguments[0];
        const isTimedGame = ['mini-sudoku', 'zip', 'crossclimb', 'queens', 'tango'].includes(gameKey);
        
        return Array.from(document.querySelectorAll('.pr-connections-leaderboard-player__container')).map(el => {
            const nameElem = el.querySelector('.pr-connections-leaderboard-player__text-wrapper .text-body-medium-bold');
            const name = nameElem ? nameElem.innerText.trim() : null;
            
            const scoreElems = Array.from(el.querySelectorAll('.text-body-medium'));
            let score = null;
            
            // 1. First priority: look for a time-like format (e.g., "1:23")
            // This is mandatory for timed games if we want to avoid streaks
            for (const s of scoreElems) {
                const text = s.innerText.trim();
                if (text.includes(':') && /^\\d+:\\d{2}$/.test(text)) {
                    score = text;
                    break;
                }
            }
            
            // 2. Second priority: If not a timed game (like Pinpoint) or if time wasn't found,
            // use the element with 'ml1' class. This is usually the rightmost column (the score).
            if (!score) {
                const altScore = el.querySelector('.text-body-medium.ml1');
                if (altScore) {
                    let text = altScore.innerText.trim();
                    // Handle leading apostrophe (often seen in Pinpoint now)
                    if (text.startsWith("'")) {
                        text = text.substring(1).trim();
                    }
                    
                    // For Pinpoint, we expect 1-5. For others, we accept any integer as a fallback
                    // but we are skeptical if it's a timed game.
                    if (/^\\d+$/.test(text)) {
                        if (gameKey === 'pinpoint') {
                            score = text;
                        } else if (!isTimedGame) {
                            score = text;
                        }
                    }
                }
            }
            
            // 3. Third priority: pick the LAST element that matches a numeric pattern
            // This is a safety fallback.
            if (!score) {
                for (let i = scoreElems.length - 1; i >= 0; i--) {
                    let text = scoreElems[i].innerText.trim();
                    if (text.startsWith("'")) {
                        text = text.substring(1).trim();
                    }
                    
                    if (/^\\d+(?::\\d{2})?$/.test(text)) {
                        // If it's a timed game and we only found an integer, it's likely a streak
                        if (isTimedGame && !text.includes(':')) continue;
                        score = text;
                        break;
                    }
                }
            }
            
            return {name, score};
        });
        """
        raw_players = self.driver.execute_script(script, game_key)
        
        players = []
        for p in raw_players:
            if p and p['name'] and p['score']:
                # Final cleanup in Python to ensure no persistent artifacts like leading apostrophes
                score = p['score'].strip()
                score = re.sub(r"^['\"`]+", "", score).strip()
                
                if score in ["-", "–", "-:--"] or not score:
                    continue
                p['score'] = score
                p['rank'] = len(players) + 1
                players.append(p)
                print(f"    {p['rank']}. {p['name']} - {p['score']}")
        
        return players
    
    def scrape_all_games(self):
        print("\n" + "="*60)
        print("Starting leaderboard scraping (Sequential)")
        print("="*60)
        
        for game_key, game_name in GAMES.items():
            print(f"\nScraping {game_name}...")
            
            # Step 1: Mandatory visit to results page for average AND user result
            results_url = f"https://www.linkedin.com/games/{game_key}/results/"
            print(f"  Fetching results from: {results_url}")
            user_score_from_results = None
            try:
                self.driver.get(results_url)
                WebDriverWait(self.driver, 10).until(
                    lambda d: d.execute_script('return document.readyState') == 'complete'
                )
                
                # Extract average
                avg = self.extract_average()
                if avg:
                    print(f"    Found average: {avg}")
                    self.all_averages[game_name] = avg
                
                # Extract user's own score as fallback
                user_score_from_results = self.extract_user_result(game_key)
                if user_score_from_results:
                    print(f"    Found user score: {user_score_from_results}")
            except Exception as e:
                print(f"    Error fetching results page: {e}")

            # Step 2: Mandatory visit to connections leaderboard for all players
            leaderboard_url = BASE_URL.format(game=game_key)
            print(f"  Fetching full leaderboard from: {leaderboard_url}")
            try:
                self.driver.get(leaderboard_url)
                WebDriverWait(self.driver, 10).until(
                    lambda d: d.execute_script('return document.readyState') == 'complete'
                )
                data = self.extract_leaderboard_data(game_key, game_name)
                
                # Ensure "You" are in the data. Prefer leaderboard score if found.
                you_in_leaderboard = next((p for p in data if p['name'].lower() == 'you'), None)
                
                if you_in_leaderboard:
                    if user_score_from_results and you_in_leaderboard['score'] != user_score_from_results:
                        print(f"    Note: Leaderboard score ({you_in_leaderboard['score']}) preferred over results page score ({user_score_from_results}).")
                elif user_score_from_results:
                    print(f"    Adding 'You' from results page: {user_score_from_results}")
                    data.append({'name': 'You', 'score': user_score_from_results, 'rank': 0})

                if data:
                    self.all_leaderboard_data[game_name] = data
                    print(f"    Successfully scraped data for {game_name}.")
                else:
                    print(f"    ⚠ No leaderboard data found for {game_name}.")
            except Exception as e:
                print(f"    Error fetching leaderboard: {e}")
            
            # Small delay between games to avoid rate-limiting
            time.sleep(1.5)
    
    def export_to_excel(self, _=None):
        # 1. Determine custom date (1:30 PM threshold)
        now = datetime.datetime.now()
        if now.hour < 13 or (now.hour == 13 and now.minute < 30):
            report_date = now - datetime.timedelta(days=1)
        else:
            report_date = now
        
        date_str = report_date.strftime('%Y-%m-%d')
        month_str = report_date.strftime('%b-%Y') # e.g., Feb-2026
        monthly_file = f"linkedin_leaderboards_{month_str}.xlsx"
        
        print(f"\nUpdating {monthly_file} for {date_str}...")
        
        # 2. Load existing sheets if file exists
        all_sheets = {}
        if os.path.exists(monthly_file):
            try:
                with pd.ExcelFile(monthly_file) as xls:
                    for s_name in xls.sheet_names:
                        all_sheets[s_name] = pd.read_excel(xls, sheet_name=s_name)
            except Exception as e:
                print(f"Error reading {monthly_file}: {e}. Starting fresh.")

        # 3. Update data for each game
        for game_name, players in self.all_leaderboard_data.items():
            if not players:
                continue
                
            new_records = []
            
            # Add Average if found
            if game_name in self.all_averages:
                new_records.append({
                    'Player Name': 'Average',
                    date_str: self.all_averages[game_name]
                })

            for p in players:
                new_records.append({
                    'Player Name': p['name'].strip(),
                    date_str: p['score']
                })
            
            new_df = pd.DataFrame(new_records)

            if game_name in all_sheets:
                df = all_sheets[game_name]
                # Clean up existing data
                df['Player Name'] = df['Player Name'].astype(str).str.strip()
                
                # If the date column already exists, drop it to overwrite
                if date_str in df.columns:
                    df = df.drop(columns=[date_str])
                
                # Merge on Player Name
                updated_df = pd.merge(df, new_df, on='Player Name', how='outer')
            else:
                updated_df = new_df
            
            all_sheets[game_name] = updated_df

        if not all_sheets:
            print("No data to export.")
            return

        # 4. Write all sheets back to the monthly file
        with pd.ExcelWriter(monthly_file, engine='openpyxl') as writer:
            # Sort games by name
            for g_name in sorted(all_sheets.keys()):
                df = all_sheets[g_name]
                
                # Reorder columns: Player Name, then sorted dates
                fixed_cols = ['Player Name']
                date_cols = sorted([c for c in df.columns if c not in fixed_cols])
                df = df[fixed_cols + date_cols]
                
                # Sort rows: "Average" first, then "You", then others alphabetically
                df['is_average'] = df['Player Name'].apply(lambda x: str(x).strip().lower() == 'average')
                df['is_you'] = df['Player Name'].apply(lambda x: str(x).strip().lower() == 'you')
                
                df = df.sort_values(
                    by=['is_average', 'is_you', 'Player Name'], 
                    ascending=[False, False, True]
                ).drop(columns=['is_average', 'is_you'])
                
                df.to_excel(writer, sheet_name=g_name, index=False)
                
                # Styling
                worksheet = writer.sheets[g_name]
                header_fill = PatternFill(start_color="0066CC", end_color="0066CC", fill_type="solid")
                header_font = Font(bold=True, color="FFFFFF")
                
                for cell in worksheet[1]:
                    cell.fill = header_fill
                    cell.font = header_font
                    cell.alignment = Alignment(horizontal='center')
                
                from openpyxl.utils import get_column_letter
                worksheet.column_dimensions['A'].width = 35 # Player Name
                for i in range(2, worksheet.max_column + 1):
                    col_letter = get_column_letter(i)
                    worksheet.column_dimensions[col_letter].width = 15

        print(f"✓ Excel file updated: {monthly_file}")
    
    def run(self, output_file='linkedin_leaderboards.xlsx'):
        try:
            self.setup_driver()
            self.driver.get("https://www.linkedin.com")
            if not self.wait_for_login():
                print("\nLogin failed or timed out. Exiting.")
                return
            self.scrape_all_games()
            self.export_to_excel(output_file)
        except KeyboardInterrupt:
            print("\nScript interrupted by user.")
        except Exception as e:
            print(f"\nAn error occurred: {e}")
        finally:
            if self.driver:
                print("\nClosing browser...")
                self.driver.quit()

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--browser-path', type=str)
    parser.add_argument('--output', type=str, default='linkedin_leaderboards.xlsx')
    args = parser.parse_args()
    
    scraper = LinkedInLeaderboardScraper(browser_executable_path=args.browser_path)
    scraper.run(output_file=args.output)

if __name__ == "__main__":
    main()
