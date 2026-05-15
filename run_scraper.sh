#!/bin/bash
# Continuous scraper runner - processes companies in small batches
# Each batch is a separate process to avoid crashes
cd /home/z/my-project
export DISPLAY=:99

BATCH_SIZE=3
TOTAL_COMPANIES=1024
MAX_BATCHES=200  # Safety limit

echo "Starting continuous scraper (batch size: $BATCH_SIZE)"

for i in $(seq 1 $MAX_BATCHES); do
    echo ""
    echo "=== BATCH $i ==="
    
    # Run one batch
    timeout 300 python3 -u job_application_scraper_v5.py $BATCH_SIZE 2>&1
    
    EXIT_CODE=$?
    echo "Batch exit code: $EXIT_CODE"
    
    # Check progress
    if [ -f job_scraper_v5_progress.json ]; then
        PROCESSED=$(python3 -c "import json; print(len(json.load(open('job_scraper_v5_progress.json'))['processed']))" 2>/dev/null || echo "?")
        JOBS=$(python3 -c "import json; print(len(json.load(open('job_scraper_v5_progress.json'))['jobs']))" 2>/dev/null || echo "?")
        echo "Progress: $PROCESSED companies, $JOBS jobs"
        
        # Stop if we've processed enough
        if [ "$PROCESSED" != "?" ] && [ "$PROCESSED" -ge "$TOTAL_COMPANIES" ]; then
            echo "All companies processed!"
            break
        fi
    fi
    
    # Brief pause between batches
    sleep 2
done

echo "Scraper runner finished."
