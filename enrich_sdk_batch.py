#!/usr/bin/env python3
"""
Website enrichment using z-ai-web-dev-sdk for web search.
Calls Node.js subprocess for each search query.
"""
import json, os, re, sys, time, hashlib, sqlite3, subprocess
from datetime import datetime
import requests

sys.stdout.reconfigure(line_buffering=True)

DB_PATH = "/home/z/my-project/download/it_employers.db"
HTML_DIR = "/home/z/my-project/download/it_employer_html"
GIT_DIR = "/home/z/my-project"
COMMIT_EVERY = 10

SESSION = requests.Session()
SESSION.headers.update({"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"})
SESSION.verify = False
import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

def sanitize_filename(name, max_len=80):
    name = re.sub(r'[äöüßÄÖÜ]', lambda m: {'ä':'ae','ö':'oe','ü':'ue','ß':'ss','Ä':'Ae','Ö':'Oe','Ü':'Ue'}[m.group()], name)
    name = re.sub(r'[^a-zA-Z0-9._-]', '_', name)
    name = re.sub(r'_+', '_', name).strip('_')
    return name[:max_len]

def search_website_sdk(company_name, city):
    """Use z-ai-web-dev-sdk via Node.js subprocess to search for website."""
    try:
        script = f'''
const ZAI = require('z-ai-web-dev-sdk').default;
async function main() {{
    const zai = await ZAI.create();
    const result = await zai.functions.invoke('web_search', {{
        query: "{company_name} {city} website",
        num: 5
    }});
    const skip = ['facebook.com','linkedin.com','instagram.com','twitter.com','youtube.com',
        'wikipedia.org','gelbeseiten.de','dasoertliche.de','meinestadt.de','kununu.com',
        'yelp.de','google.com','xing.com','northdata.de','kompass.com','firmenwissen.de'];
    for (const r of result) {{
        const u = r.url.toLowerCase();
        if (!skip.some(d => u.includes(d))) {{
            console.log(r.url);
            return;
        }}
    }}
    if (result.length > 0) console.log(result[0].url);
    else console.log('');
}}
main().catch(e => console.error(''));
'''
        result = subprocess.run(['node', '-e', script], capture_output=True, text=True, timeout=15)
        url = result.stdout.strip()
        if url and url.startswith('http'):
            return url
    except:
        pass
    return None

def download_html(url, timeout=8):
    try:
        r = SESSION.get(url, timeout=timeout, allow_redirects=True)
        if r.status_code == 200 and 'text/' in r.headers.get('content-type', ''):
            return r.text
    except:
        pass
    return None

def save_html_file(content, name, suffix=""):
    if not content: return None, None
    fn = f"{sanitize_filename(name)}_{suffix}.html" if suffix else f"{sanitize_filename(name)}.html"
    fp = os.path.join(HTML_DIR, fn)
    os.makedirs(HTML_DIR, exist_ok=True)
    with open(fp, 'w', encoding='utf-8', errors='replace') as f: f.write(content)
    return fp, hashlib.sha256(content.encode()).hexdigest()[:16]

def make_abs(base, rel):
    if not rel: return None
    if rel.startswith('http'): return rel
    if rel.startswith('//'): return 'https:' + rel
    if rel.startswith('/'):
        from urllib.parse import urlparse
        p = urlparse(base)
        return f"{p.scheme}://{p.netloc}{rel}"
    return f"{base.rstrip('/')}/{rel.lstrip('/')}"

def extract_emails(html):
    emails = re.findall(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}', html)
    skip = ['example.com','sentry','webpack','noreply','wixpress','github','googleapis']
    return list(dict.fromkeys(e for e in emails if not any(s in e.lower() for s in skip)))[:3]

def extract_links(html, base_url):
    imp_url = None; kar_url = None
    soup = BeautifulSoup(html, 'html.parser')
    for a in soup.find_all('a', href=True):
        h = a['href'].lower(); t = a.get_text(strip=True).lower()
        if not imp_url and ('impressum' in h or 'impressum' in t): imp_url = a['href']
        if not kar_url and any(w in h for w in ['karriere','career','jobs','stellenangebote','bewerbung']): kar_url = a['href']
    return imp_url, kar_url

from bs4 import BeautifulSoup

def main():
    print("=== Website Enrichment via SDK ===")
    conn = sqlite3.connect(DB_PATH)
    
    # Get IT employers without website
    rows = conn.execute("""SELECT id, firmenname, ort, hat_it_jobs, it_relevanz_score 
        FROM employers WHERE website_url = '' AND hat_it_jobs = 1
        ORDER BY it_relevanz_score DESC""").fetchall()
    print(f"IT-Arbeitgeber ohne Website: {len(rows)}")
    
    total = len(rows); processed = 0; found = 0; batch_commit = 0
    
    for row in rows:
        emp_id, name, ort, has_it, score = row
        processed += 1
        
        # Search via SDK
        website = search_website_sdk(name, ort)
        
        if website:
            found += 1
            html_path = None; html_hash = None; imp_path = None; kar_path = None
            imp_url = None; kar_url = None; email = ''
            
            html = download_html(website)
            if html:
                html_path, html_hash = save_html_file(html, name)
                
                emails = extract_emails(html)
                if emails: email = emails[0]
                
                imp_rel, kar_rel = extract_links(html, website)
                if imp_rel:
                    imp_url = make_abs(website, imp_rel)
                    imp_html = download_html(imp_url, timeout=6)
                    if imp_html:
                        imp_path, _ = save_html_file(imp_html, name, "impressum")
                        imp_emails = extract_emails(imp_html)
                        if imp_emails and not email: email = imp_emails[0]
                
                if kar_rel:
                    kar_url = make_abs(website, kar_rel)
                    kar_html = download_html(kar_url, timeout=6)
                    if kar_html:
                        kar_path, _ = save_html_file(kar_html, name, "karriere")
            
            conn.execute("""UPDATE employers SET website_url=?, website_html_path=?, website_html_hash=?,
                impressum_url=?, impressum_html_path=?, karriere_url=?, karriere_html_path=?,
                email=COALESCE(NULLIF(email,''),?), scrape_status='enriched',
                updated_at=CURRENT_TIMESTAMP WHERE id=?""",
                [website, html_path or '', html_hash or '', imp_url or '', imp_path or '',
                 kar_url or '', kar_path or '', email, emp_id])
            conn.commit()
            
            print(f"  [{processed}/{total}] ✓IT Sc:{score:5.1f} | {name[:40]:40} | {website[:40]}")
        else:
            conn.execute("UPDATE employers SET scrape_status='no_website', updated_at=CURRENT_TIMESTAMP WHERE id=?", [emp_id])
            conn.commit()
            print(f"  [{processed}/{total}]    Sc:{score:5.1f} | {name[:40]:40} | -")
        
        batch_commit += 1
        time.sleep(0.3)
        
        if batch_commit >= COMMIT_EVERY:
            # Rebuild Excel and commit
            try:
                subprocess.run(['python3', '/home/z/my-project/build_it_excel.py'], capture_output=True, timeout=30)
            except: pass
            try:
                subprocess.run(['git','add','-A'], cwd=GIT_DIR, capture_output=True, timeout=30)
                msg = f"Website-SDK: {processed}/{total} | {found} found ({datetime.now().strftime('%H:%M')})"
                r = subprocess.run(['git','commit','-m',msg], cwd=GIT_DIR, capture_output=True, timeout=30)
                if r.returncode == 0:
                    subprocess.run(['git','push','origin','main'], cwd=GIT_DIR, capture_output=True, timeout=60)
                    print(f"  >>> GIT PUSH ({found} Websites) <<<")
            except Exception as e: print(f"  GIT: {e}")
            batch_commit = 0
    
    # Final
    try:
        subprocess.run(['python3', '/home/z/my-project/build_it_excel.py'], capture_output=True, timeout=30)
    except: pass
    try:
        subprocess.run(['git','add','-A'], cwd=GIT_DIR, capture_output=True, timeout=30)
        subprocess.run(['git','commit','-m',f"Website-SDK FERTIG: {found}/{total}"], cwd=GIT_DIR, capture_output=True, timeout=30)
        subprocess.run(['git','push','origin','main'], cwd=GIT_DIR, capture_output=True, timeout=60)
    except: pass
    
    conn.close()
    print(f"\nFERTIG! {found}/{total} Websites für IT-Arbeitgeber gefunden")

if __name__ == '__main__':
    main()
