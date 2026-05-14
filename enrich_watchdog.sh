#!/bin/bash
cd /home/z/my-project
while true; do
    REMAINING=$(python3 -c "
import json
with open('/home/z/my-project/gmaps_erlangen_results.json') as f: r = json.load(f)
with open('/home/z/my-project/website_enrichment.json') as f: e = json.load(f)
cities = ['Bamberg','Erlangen','Nürnberg']
need = [x for x in r if x.get('city') in cities and x.get('name','').strip() and x['name'].strip() not in e]
print(len(need))
" 2>/dev/null || echo "999")
    
    echo "[$(date +%H:%M:%S)] Remaining: $REMAINING"
    
    if [ "$REMAINING" = "0" ]; then
        echo "ALL DONE!"
        break
    fi
    
    # Run a batch
    timeout 120 python3 enrich_v2.py 0 25 2>&1 | tail -5
    echo "[$(date +%H:%M:%S)] Batch done, sleeping 5s..."
    sleep 5
done
