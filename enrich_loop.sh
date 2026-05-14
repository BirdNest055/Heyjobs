#!/bin/bash
# Process IT employers without websites, 10 at a time
DB="/home/z/my-project/download/it_employers.db"
GIT="/home/z/my-project"

COUNT=0
FOUND=0

while true; do
    # Get next employer without website
    ROW=$(python3 -c "
import sqlite3,json
conn=sqlite3.connect('$DB')
row=conn.execute(\"SELECT id, firmenname, ort FROM employers WHERE website_url='' AND hat_it_jobs=1 ORDER BY it_relevanz_score DESC LIMIT 1\").fetchone()
if row: print(json.dumps(row))
else: print('')
conn.close()
")
    
    [ -z "$ROW" ] || [ "$ROW" = "" ] && break
    
    EMP_ID=$(echo $ROW | python3 -c "import json,sys; print(json.load(sys.stdin)[0])")
    NAME=$(echo $ROW | python3 -c "import json,sys; print(json.load(sys.stdin)[1])")
    ORT=$(echo $ROW | python3 -c "import json,sys; print(json.load(sys.stdin)[2])")
    
    COUNT=$((COUNT + 1))
    
    # Search via Node.js SDK
    WEBSITE=$(timeout 20 node -e "
const Z=require('z-ai-web-dev-sdk').default;
async function m(){
    const z=await Z.create();
    const r=await z.functions.invoke('web_search',{query:'${NAME} ${ORT} website',num:5});
    const skip=['facebook.com','linkedin.com','wikipedia.org','gelbeseiten.de','kunumu.com','google.com','northdata.de','firmenwissen.de'];
    for(const x of r){if(!skip.some(d=>x.url.toLowerCase().includes(d))){console.log(x.url);return;}}
    if(r.length>0)console.log(r[0].url);
}
m().catch(()=>console.log(''));
" 2>/dev/null)
    
    if [ -n "$WEBSITE" ] && [[ "$WEBSITE" == http* ]]; then
        FOUND=$((FOUND + 1))
        # Update DB
        python3 -c "
import sqlite3
conn=sqlite3.connect('$DB')
conn.execute(\"UPDATE employers SET website_url = ?, scrape_status = 'enriched', updated_at = CURRENT_TIMESTAMP WHERE id = ?\", ['$WEBSITE', $EMP_ID])
conn.commit()
conn.close()
"
        echo "  [$COUNT] + ${NAME:0:40} | ${WEBSITE:0:40}"
    else
        python3 -c "
import sqlite3
conn=sqlite3.connect('$DB')
conn.execute(\"UPDATE employers SET scrape_status = 'no_website' WHERE id = ?\", [$EMP_ID])
conn.commit()
conn.close()
"
        echo "  [$COUNT]   ${NAME:0:40} | -"
    fi
    
    # Git every 10
    if [ $((COUNT % 10)) -eq 0 ]; then
        cd $GIT
        git add -A 2>/dev/null
        git commit -m "SDK-Loop: $COUNT done, $FOUND websites" 2>/dev/null
        git push origin main 2>/dev/null
        echo "  >>> GIT PUSH ($FOUND websites) <<<"
    fi
    
    sleep 0.3
done

# Final commit
cd $GIT
git add -A 2>/dev/null
git commit -m "SDK-Loop FERTIG: $FOUND/$COUNT" 2>/dev/null
git push origin main 2>/dev/null
echo "FERTIG! $FOUND/$COUNT"
