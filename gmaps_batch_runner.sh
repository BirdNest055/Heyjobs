#!/bin/bash
# Runs the detail scraper one city at a time
cd /home/z/my-project

while true; do
    # Check if there are cities left to process
    python3 -c "
import json
with open('gmaps_detail_progress.json') as f:
    d = json.load(f)
cities_done = set(d.get('cities_done', []))
all_cities = ['Erlangen', 'Nürnberg', 'Fürth', 'Bamberg', 'Herzogenaurach', 'Schwabach', 'Forchheim', 'Lauf an der Pegnitz', 'Hersbruck', 'Neustadt an der Aisch', 'Eckental', 'Ebermannstadt', 'Roth', 'Altdorf bei Nürnberg', 'Zirndorf', 'Strullendorf', 'Höchstadt an der Aisch']
remaining = [c for c in all_cities if c not in cities_done]
if remaining:
    print(f'REMAINING: {len(remaining)} cities - next: {remaining[0]}')
    exit(0)
else:
    print('ALL DONE')
    exit(1)
"
    if [ $? -ne 0 ]; then
        echo "All cities processed!"
        break
    fi
    
    echo "=== Running next batch ==="
    # Kill any leftover processes
    pkill -9 Xvfb 2>/dev/null
    pkill -9 chromium 2>/dev/null
    pkill -9 chrome 2>/dev/null
    sleep 2
    
    # Run the scraper (will process one city then exit when it hits issues)
    timeout 600 python3 gmaps_detail_rescraper.py 2>&1 | tee -a gmaps_detail_log.txt
    
    echo "Batch completed, sleeping before next..."
    sleep 5
done
echo "=== All done! ==="
