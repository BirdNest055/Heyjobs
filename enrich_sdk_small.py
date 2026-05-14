#!/usr/bin/env python3
"""Small batch SDK enrichment - processes 20 at a time, then commits."""
import json, os, re, sys, time, hashlib, sqlite3, subprocess
from datetime import datetime
import requests
from bs4 import BeautifulSoup

sys.stdout.reconfigure(line_buffering=True)

DB_PATH = "/home/z/my-project/download/it_employers.db"
HTML_DIR = "/home/z/my-project/download/it_employer_html"
GIT_DIR = "/home/z/my-project"
BATCH = 20

SESSION = requests.Session()
SESSION.headers.update({"User-Agent": "Mozilla/5.0"})
SESSION.verify = False
import urllib3; urllib3.disable_warnings()

def sanitize_fn(n, m=80):
    n = re.sub(r'[äöüßÄÖÜ]', lambda x: {'ä':'ae','ö':'oe','ü':'ue','ß':'ss','Ä':'Ae','Ö':'Oe','Ü':'Ue'}[x.group()], n)
    return re.sub(r'[^a-zA-Z0-9._-]','_',n).replace('_+','_').strip('_')[:m]

def search_ws(name, city):
    try:
        r = subprocess.run(['node','-e',f'''
const Z=require("z-ai-web-dev-sdk").default;
async function m(){{const z=await Z.create();const r=await z.functions.invoke("web_search",{{query:"{name} {city} website",num:5}});const s=["facebook.com","linkedin.com","wikipedia.org","gelbeseiten.de","kununu.com","northdata.de","firmenwissen.de","google.com"];for(const x of r){{if(!s.some(d=>x.url.toLowerCase().includes(d))){{console.log(x.url);return;}}}}if(r.length>0)console.log(r[0].url);else console.log("");}}m().catch(()=>console.log(""));
'''], capture_output=True, text=True, timeout=15)
        url = r.stdout.strip()
        if url and url.startswith('http'): return url
    except: pass
    return None

def dl(url, t=8):
    try:
        r = SESSION.get(url, timeout=t, allow_redirects=True)
        if r.status_code==200 and 'text/' in r.headers.get('content-type',''): return r.text
    except: pass
    return None

def save_h(c, n, s=""):
    if not c: return None,None
    fn = f"{sanitize_fn(n)}_{s}.html" if s else f"{sanitize_fn(n)}.html"
    fp = os.path.join(HTML_DIR, fn)
    os.makedirs(HTML_DIR, exist_ok=True)
    with open(fp,'w',encoding='utf-8',errors='replace') as f: f.write(c)
    return fp, hashlib.sha256(c.encode()).hexdigest()[:16]

def mk_abs(b, r):
    if not r: return None
    if r.startswith('http'): return r
    if r.startswith('//'): return 'https:'+r
    if r.startswith('/'):
        from urllib.parse import urlparse
        p=urlparse(b); return f"{p.scheme}://{p.netloc}{r}"
    return f"{b.rstrip('/')}/{r.lstrip('/')}"

def ext_email(h):
    es = re.findall(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}',h)
    sk = ['example.com','sentry','webpack','noreply','github']
    return list(dict.fromkeys(e for e in es if not any(s in e.lower() for s in sk)))[:3]

def ext_links(h):
    s = BeautifulSoup(h,'html.parser'); imp=None; kar=None
    for a in s.find_all('a',href=True):
        hr=a['href'].lower(); t=a.get_text(strip=True).lower()
        if not imp and ('impressum' in hr or 'impressum' in t): imp=a['href']
        if not kar and any(w in hr for w in ['karriere','career','jobs','stellenangebote','bewerbung']): kar=a['href']
    return imp, kar

def main():
    print("=== SDK Enrichment (small batches) ===")
    conn = sqlite3.connect(DB_PATH)
    rows = conn.execute("""SELECT id, firmenname, ort, it_relevanz_score 
        FROM employers WHERE website_url='' AND hat_it_jobs=1
        ORDER BY it_relevanz_score DESC""").fetchall()
    print(f"IT-Arbeitgeber ohne Website: {len(rows)}")
    
    total=len(rows); proc=0; found=0; bc=0
    for eid,name,ort,score in rows:
        proc += 1
        ws = search_ws(name, ort)
        
        if ws:
            found += 1
            h=dl(ws); hp=hh=ip=iup=kp=ku=None; em=''
            if h:
                hp,hh=save_h(h,name)
                es=ext_email(h)
                if es: em=es[0]
                ir,kr=ext_links(h)
                if ir:
                    iua=mk_abs(ws,ir)
                    if iua: iup=iua; ih=dl(iua,6)
                    if ih: ip,_=save_h(ih,name,"impressum"); ie=ext_email(ih)
                    if ie and not em: em=ie[0]
                if kr:
                    kua=mk_abs(ws,kr)
                    if kua: ku=kua; kh=dl(kua,6)
                    if kh: kp,_=save_h(kh,name,"karriere")
            conn.execute("""UPDATE employers SET website_url=?,website_html_path=?,website_html_hash=?,
                impressum_url=?,impressum_html_path=?,karriere_url=?,karriere_html_path=?,
                email=COALESCE(NULLIF(email,''),?),scrape_status='enriched',
                updated_at=CURRENT_TIMESTAMP WHERE id=?""",
                [ws,hp or '',hh or '',iup or '',ip or '',ku or '',kp or '',em,eid])
            conn.commit()
            print(f"  [{proc}/{total}] ✓ Sc:{score:5.1f} | {name[:38]:38} | {ws[:35]}")
        else:
            conn.execute("UPDATE employers SET scrape_status='no_website',updated_at=CURRENT_TIMESTAMP WHERE id=?",[eid])
            conn.commit()
            print(f"  [{proc}/{total}]   Sc:{score:5.1f} | {name[:38]:38} | -")
        
        bc += 1; time.sleep(0.2)
        if bc >= BATCH:
            try:
                subprocess.run(['git','add','-A'],cwd=GIT_DIR,capture_output=True,timeout=30)
                msg=f"SDK-Enrich: {proc}/{total} | {found} ws ({datetime.now().strftime('%H:%M')})"
                subprocess.run(['git','commit','-m',msg],cwd=GIT_DIR,capture_output=True,timeout=30)
                subprocess.run(['git','push','origin','main'],cwd=GIT_DIR,capture_output=True,timeout=60)
                print(f"  >>> PUSH ({found}) <<<")
            except: pass
            bc = 0
    
    # Final
    try:
        subprocess.run(['git','add','-A'],cwd=GIT_DIR,capture_output=True,timeout=30)
        subprocess.run(['git','commit','-m',f"SDK-Enrich FERTIG: {found}/{total}"],cwd=GIT_DIR,capture_output=True,timeout=30)
        subprocess.run(['git','push','origin','main'],cwd=GIT_DIR,capture_output=True,timeout=60)
    except: pass
    conn.close()
    print(f"\nFERTIG! {found}/{total}")

if __name__=='__main__':
    main()
