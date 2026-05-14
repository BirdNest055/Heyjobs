#!/bin/bash
cd /home/z/my-project
for i in $(seq 0 20 1140); do
    echo "[$(date +%H:%M:%S)] Batch starting at $i"
    timeout 300 python3 enrich_v2.py $i 20 2>&1 | tail -3
    
    # Check remaining
    REMAINING=$(python3 -c "
import json
with open('/home/z/my-project/gmaps_erlangen_results.json') as f: r = json.load(f)
with open('/home/z/my-project/website_enrichment.json') as f: e = json.load(f)
cities = ['Bamberg','Erlangen','Nürnberg']
need = [x for x in r if x.get('city') in cities and x.get('name','').strip() and x['name'].strip() not in e]
print(len(need))
")
    echo "[$(date +%H:%M:%S)] Remaining: $REMAINING"
    
    if [ "$REMAINING" -eq 0 ]; then
        echo "[$(date +%H:%M:%S)] ALL DONE!"
        break
    fi
    
    sleep 2
done
echo "[$(date +%H:%M:%S)] Loop finished"
