#!/bin/bash

# Use PowerShell to show a popup message box
# 4 = Yes/No buttons, 32 = Question icon
# Use PowerShell to show a popup message box and trim whitespace/carriage returns
CHOICE=$(powershell.exe -Command "(New-Object -ComObject WScript.Shell).Popup('Do you want to run the LinkedIn Leaderboard Scraper now?', 0, 'LinkedIn Scraper Task', 4 + 32)" | tr -d '\r' | xargs)

# 6 is the result code for "Yes"
if [ "$CHOICE" = "6" ]; then
    echo "Starting scraper..."
    cd "$(dirname "$0")"
    ./run
else
    echo "Scraper skipped by user ($CHOICE)."
fi
