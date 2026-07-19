# LinkedIn Games Leaderboard Scraper

Automates your personal browser to fetch leaderboard results from LinkedIn games and exports them to Excel.

## Features

- ✅ Scrapes leaderboards from all LinkedIn games (Mini Sudoku, Zip, Pinpoint, Crossclimb, Queens, Tango)
- ✅ Extracts player names and scores
- ✅ Exports to Excel with one sheet per game
- ✅ Uses your personal browser (Chrome/Chromium)
- ✅ Formatted Excel output with headers

## Installation

1. Install Python dependencies:
```bash
pip install -r requirements.txt
```

2. You'll also need ChromeDriver. Install it via:
```bash
# On Linux/Ubuntu
sudo apt-get install chromium-chromedriver

# Or download manually from:
# https://chromedriver.chromium.org/downloads
```

## Usage

### Basic Usage (with system Chrome)
```bash
python linkedin_leaderboard_scraper.py
```

### With Custom Browser Path
```bash
python linkedin_leaderboard_scraper.py --browser-path /path/to/chrome
```

### Custom Output File
```bash
python linkedin_leaderboard_scraper.py --output my_leaderboards.xlsx
```

### All Options
```bash
python linkedin_leaderboard_scraper.py \
  --browser-path /usr/bin/google-chrome \
  --output leaderboards_2026.xlsx
```

## How It Works

1. Opens your Chrome browser
2. Navigates to LinkedIn
3. **Waits for you to manually log in** (you'll need to log in once)
4. Sequentially visits each game's leaderboard page
5. Scrapes player names and scores
6. Exports everything to Excel

## Output Format

The Excel file will contain one sheet per game with columns:
- **Rank**: Player ranking (1, 2, 3, ...)
- **Player Name**: Name of the player
- **Score**: Time or score (e.g., "0:37")

## Common Browser Paths

### Linux
- `/usr/bin/google-chrome`
- `/usr/bin/chromium-browser`
- `/usr/bin/chromium`

### Windows
- `C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe`
- `C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe`

### macOS
- `/Applications/Google Chrome.app/Contents/MacOS/Google Chrome`

## Troubleshooting

**ChromeDriver version mismatch:**
Make sure ChromeDriver version matches your Chrome version.

**Login timeout:**
If you need more time to log in, the script waits for you to press ENTER after logging in manually.

**No data scraped:**
- Make sure you're logged into LinkedIn
- Check if LinkedIn's HTML structure has changed
- The script looks for elements with class `pr-connections-leaderboard-player__container`

## Notes

- The script includes polite delays between requests
- You only need to log in once per session
- The browser will close automatically after scraping completes
