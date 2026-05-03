#!/bin/bash
export DISPLAY=:99
cd /home/z/my-project

# Kill old processes
pkill -f gmaps_town_worker 2>/dev/null
pkill Xvfb 2>/dev/null
sleep 1

# Start Xvfb
Xvfb :99 -screen 0 1920x1080x24 >/dev/null 2>&1 &
sleep 2

echo "[$(date +%H:%M:%S)] Starting town-by-town scraper loop"

# Process 3 towns per iteration, then restart browser
for i in $(seq 1 15); do
    echo "[$(date +%H:%M:%S)] === Batch $i ==="
    timeout 300 python3 /home/z/my-project/gmaps_robust_continue.py 3 2>&1 | tail -20
    EXIT=$?
    echo "[$(date +%H:%M:%S)] Batch $i exit code: $EXIT"
    
    # Check how many remaining
    REMAINING=$(python3 -c "
import json
with open('/home/z/my-project/gmaps_erlangen_progress.json') as f:
    p = json.load(f)
with open('/home/z/my-project/gmaps_erlangen_towns.json') as f:
    t = json.load(f)
done = set(cs.split('|')[0] for cs in p['completed_searches'])
left = [x for x in t if x['name'] not in done]
print(len(left))
")
    echo "[$(date +%H:%M:%S)] Remaining towns: $REMAINING"
    
    if [ "$REMAINING" -eq 0 ]; then
        echo "[$(date +%H:%M:%S)] ALL TOWNS DONE!"
        break
    fi
    
    # Restart Xvfb between batches
    pkill Xvfb 2>/dev/null
    sleep 2
    Xvfb :99 -screen 0 1920x1080x24 >/dev/null 2>&1 &
    sleep 2
done

echo "[$(date +%H:%M:%S)] Loop finished"
