#!/bin/bash
cd /home/z/my-project
export PYTHONUNBUFFERED=1
export DISPLAY=:99

MAX_RETRIES=50
RETRY=0

while [ $RETRY -lt $MAX_RETRIES ]; do
    echo "=== Starting scraper (attempt $((RETRY+1))) ===" >> scrape_wrapper_log.txt
    date >> scrape_wrapper_log.txt
    
    timeout 600 python3 -u junior_jobs_scraper_v4.py >> scrape_wrapper_log.txt 2>&1
    EXIT_CODE=$?
    
    echo "=== Scraper exited with code $EXIT_CODE ===" >> scrape_wrapper_log.txt
    date >> scrape_wrapper_log.txt
    
    if [ $EXIT_CODE -eq 0 ]; then
        echo "Scraper completed successfully!" >> scrape_wrapper_log.txt
        break
    fi
    
    # Check if there are still employers to process
    python3 -c "
import json
p = json.load(open('download/scrape_v3_progress.json'))
remaining = 254 - len(p.get('completed', []))
print(f'Remaining: {remaining}')
if remaining <= 0:
    exit(1)  # Signal to stop
" 2>/dev/null
    
    if [ $? -ne 0 ]; then
        echo "All employers processed!" >> scrape_wrapper_log.txt
        break
    fi
    
    RETRY=$((RETRY+1))
    echo "Waiting 10s before retry..." >> scrape_wrapper_log.txt
    sleep 10
done

echo "=== Wrapper finished ===" >> scrape_wrapper_log.txt
