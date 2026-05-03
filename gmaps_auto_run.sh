#!/usr/bin/env bash
# Auto-run all towns in batches of 5, with proper cleanup
cd /home/z/my-project

# Check how many searches are done
COMPLETED=$(python3 -c "
import json
p = json.load(open('gmaps_erlangen_progress.json'))
print(len(p['completed_searches']))
")
TOTAL_SEARCHES=1040  # estimated total

echo "Completed searches: $COMPLETED / ~$TOTAL_SEARCHES"

# Run in batches of 5 towns
for START in $(seq 15 5 170); do
    # Check if process is already running
    if pgrep -f gmaps_lightweight > /dev/null; then
        echo "Already running, waiting..."
        wait
    fi
    
    echo "=== Starting batch: towns $START to $((START+4)) ==="
    python3 gmaps_lightweight.py $START 5 2>&1
    
    # Brief pause between batches
    sleep 5
    
    # Check total
    TOTAL=$(python3 -c "import json; d=json.load(open('gmaps_erlangen_results.json')); print(len(d))")
    echo "=== Total so far: $TOTAL businesses ==="
done

echo "All done!"
