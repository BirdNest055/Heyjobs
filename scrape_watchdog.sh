#!/bin/bash
cd /home/z/my-project
export PYTHONUNBUFFERED=1
export DISPLAY=:99

while true; do
    # Check if there are still employers to process
    REMAINING=$(python3 -c "
import json
p = json.load(open('download/scrape_v3_progress.json'))
completed = len(p.get('completed', []))
total = 254  # known count
print(total - completed)
" 2>/dev/null)
    
    if [ -z "$REMAINING" ] || [ "$REMAINING" -le 0 ]; then
        echo "All employers processed!"
        break
    fi
    
    echo "=== Starting scraper (remaining: $REMAINING) ==="  >> scrape_watchdog_log.txt
    date >> scrape_watchdog_log.txt
    
    # Run with hard timeout of 300 seconds (5 minutes max per run)
    timeout 300 python3 -u junior_jobs_scraper_v4.py >> scrape_watchdog_log.txt 2>&1
    
    EXIT_CODE=$?
    echo "=== Exited with code $EXIT_CODE ===" >> scrape_watchdog_log.txt
    
    # Wait before restarting
    sleep 5
done

echo "=== Watchdog finished ===" >> scrape_watchdog_log.txt
