#!/bin/bash
cd /home/z/my-project
export PYTHONUNBUFFERED=1

MAX_ATTEMPTS=100
ATTEMPT=0

while [ $ATTEMPT -lt $MAX_ATTEMPTS ]; do
    ATTEMPT=$((ATTEMPT+1))
    
    # Check remaining
    REMAINING=$(python3 -c "
import json
p = json.load(open('download/scrape_v3_progress.json'))
total = 254
completed = len(p.get('completed', []))
print(total - completed)
" 2>/dev/null)
    
    if [ -z "$REMAINING" ] || [ "$REMAINING" -le 0 ] 2>/dev/null; then
        echo "All done!" >> scraper_robust_log.txt
        break
    fi
    
    echo "=== Attempt $ATTEMPT (remaining: $REMAINING) ===" >> scraper_robust_log.txt
    date >> scraper_robust_log.txt
    
    # Run with 120s timeout - if it hangs, kill it
    timeout 120 python3 -u junior_jobs_scraper_v4.py >> scraper_robust_log.txt 2>&1
    EXIT=$?
    
    echo "=== Exit code: $EXIT ===" >> scraper_robust_log.txt
    
    # Wait a bit before retry
    sleep 5
done

echo "=== Scraper loop finished ===" >> scraper_robust_log.txt
date >> scraper_robust_log.txt
