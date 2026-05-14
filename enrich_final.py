#!/usr/bin/env python3
"""
Final website enrichment using z-ai-web-dev-sdk via Node.js subprocess.
Processes one at a time, commits every 10.
"""
import json, os, sys, sqlite3, subprocess, time
from datetime import datetime

sys.stdout.reconfigure(line_buffering=True)

DB = "/home/z/my-project/download/it_employers.db"
GIT = "/home/z/my-project"

def search_website_node(name, city):
    """Search for website using Node.js SDK."""
    # Escape quotes in name
    safe_name = name.replace("'", "\\'").replace('"', '\\"')
    safe_city = city.replace("'", "\\'").replace('"', '\\"')
    
    script = f'''const Z=require("z-ai-web-dev-sdk").default;
async function m(){{
    const z=await Z.create();
    const r=await z.functions.invoke("web_search",{{query:"{safe_name} {safe_city} website",num:5}});
    const skip=["facebook.com","linkedin.com","wikipedia.org","gelbeseiten.de","kununu.com","google.com","northdata.de","firmenwissen.de","yelp.de","dasoertliche.de"];
    for(const x of r){{if(!skip.some(d=>x.url.toLowerCase().includes(d))){{console.log(x.url);return;}}}}
    if(r.length>0)console.log(r[0].url);
}}
m().catch(()=>console.log(""));'''
    
    try:
        result = subprocess.run(['node', '-e', script], capture_output=True, text=True, timeout=20)
        url = result.stdout.strip()
        if url and url.startswith('http'):
            return url
    except:
        pass
    return None

def main():
    print("=== Final SDK Enrichment ===")
    conn = sqlite3.connect(DB)
    
    rows = conn.execute("""SELECT id, firmenname, ort, it_relevanz_score 
        FROM employers WHERE website_url='' AND hat_it_jobs=1
        ORDER BY it_relevanz_score DESC""").fetchall()
    print(f"IT-Arbeitgeber ohne Website: {len(rows)}")
    
    total = len(rows)
    processed = 0
    found = 0
    batch = 0
    
    for emp_id, name, ort, score in rows:
        processed += 1
        
        website = search_website_node(name, ort)
        
        if website:
            found += 1
            # Escape for SQL
            safe_ws = website.replace("'", "''")
            conn.execute(f"UPDATE employers SET website_url = '{safe_ws}', scrape_status = 'enriched', updated_at = CURRENT_TIMESTAMP WHERE id = {emp_id}")
            conn.commit()
            print(f"  [{processed}/{total}] + Sc:{score:5.1f} | {name[:38]:38} | {website[:40]}")
        else:
            conn.execute(f"UPDATE employers SET scrape_status = 'no_website' WHERE id = {emp_id}")
            conn.commit()
            print(f"  [{processed}/{total}]   Sc:{score:5.1f} | {name[:38]:38} | -")
        
        batch += 1
        time.sleep(0.3)
        
        if batch >= 10:
            try:
                subprocess.run(['git', 'add', '-A'], cwd=GIT, capture_output=True, timeout=30)
                msg = f"SDK-Final: {processed}/{total} | {found} websites ({datetime.now().strftime('%H:%M')})"
                r = subprocess.run(['git', 'commit', '-m', msg], cwd=GIT, capture_output=True, timeout=30)
                if r.returncode == 0:
                    subprocess.run(['git', 'push', 'origin', 'main'], cwd=GIT, capture_output=True, timeout=60)
                    print(f"  >>> GIT PUSH ({found} websites) <<<")
            except: pass
            batch = 0
    
    # Final
    try:
        subprocess.run(['git', 'add', '-A'], cwd=GIT, capture_output=True, timeout=30)
        subprocess.run(['git', 'commit', '-m', f"SDK-Final FERTIG: {found}/{total}"], cwd=GIT, capture_output=True, timeout=30)
        subprocess.run(['git', 'push', 'origin', 'main'], cwd=GIT, capture_output=True, timeout=60)
    except: pass
    
    conn.close()
    print(f"\nFERTIG! {found}/{total} Websites gefunden")

if __name__ == '__main__':
    main()
