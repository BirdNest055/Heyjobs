#!/usr/bin/env python3
"""Website Enrichment using z-ai-web-dev-sdk for search (avoids DDG rate limits)."""
import json, re, time, os, sys, traceback, subprocess, socket
from datetime import datetime
import urllib.request, urllib.parse, urllib.error

RESULTS_FILE = '/home/z/my-project/gmaps_erlangen_results.json'
ENRICH_FILE = '/home/z/my-project/website_enrichment.json'
HTML_DIR = '/home/z/my-project/download/website_html'
LOG_FILE = '/home/z/my-project/enrich_zai_log.txt'
TARGET_CITIES = ['Bamberg', 'Erlangen', 'Nürnberg']

def log(msg):
    ts = datetime.now().strftime('%H:%M:%S')
    print(f"[{ts}] {msg}", flush=True)
    with open(LOG_FILE, 'a') as f:
        f.write(f"[{ts}] {msg}\n")

def safe_filename(name):
    name = re.sub(r'[^\w\s\-.]', '', name)
    name = re.sub(r'\s+', '_', name.strip())
    return name[:80]

def search_website_zai(name, city):
    """Use z-ai web search to find company website."""
    query = f'{name} {city} Webseite'
    try:
        result = subprocess.run(
            ['node', '-e', f'''
const ZAI = require("z-ai-web-dev-sdk").default;
async function search() {{
    const zai = await ZAI.create();
    const r = await zai.functions.invoke("web_search", {{query: {json.dumps(query)}, num: 5}});
    console.log(JSON.stringify(r));
}}
search().catch(e => console.error(JSON.stringify({{error: e.message}})));
'''],
            capture_output=True, text=True, timeout=15
        )
        data = json.loads(result.stdout.strip())
        if isinstance(data, dict) and 'error' in data:
            return None
        
        skip = ['facebook.com','instagram.com','twitter.com','linkedin.com','youtube.com',
                'wikipedia.org','maps.google','gelbeseiten.de','yelp.de','tripadvisor',
                'firmenwissen.de','northdata','kompass','wlw.de','cylex.de']
        
        for item in data:
            url = item.get('url', '')
            if not any(s in url.lower() for s in skip):
                return url
        
        # Return first result if all are skip domains
        if data:
            return data[0].get('url', '')
        
        return None
    except Exception as e:
        log(f"  z-ai search error: {str(e)[:60]}")
        return None

def url_get(url, timeout=15, retries=1):
    """HTTP GET with retries."""
    for attempt in range(retries + 1):
        try:
            req = urllib.request.Request(url, headers={
                'User-Agent': 'Mozilla/5.0 (X11; Linux x86_64; rv:124.0) Gecko/20100101 Firefox/124.0',
                'Accept': 'text/html,application/xhtml+xml,*/*',
                'Accept-Language': 'de-DE,de;q=0.9,en;q=0.5',
            })
            with urllib.request.urlopen(req, timeout=timeout) as resp:
                return resp.read(), resp.url, resp.headers.get('Content-Type','')
        except (urllib.error.URLError, socket.timeout, ConnectionError, OSError) as e:
            if attempt < retries:
                time.sleep(3)
    return None, None, None

def extract_emails(html_bytes):
    """Extract emails from HTML."""
    try: text = html_bytes.decode('utf-8', errors='replace')
    except: text = html_bytes.decode('latin-1', errors='replace')
    
    emails = set()
    for m in re.finditer(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}', text):
        e = m.group(0).lower()
        if not any(x in e for x in ['.png','.jpg','.gif','.svg','.css','.js','example.com','sentry','wixpress','googlemail']):
            emails.add(e)
    
    for m in re.finditer(r'mailto:([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})', text, re.I):
        emails.add(m.group(1).lower())
    
    skip = ['sage','wordpress','admin@localhost','test@','noreply@','webmaster@','postmaster@','donotreply@']
    emails = {e for e in emails if not any(s in e for s in skip)}
    
    pref = [e for e in emails if any(k in e for k in ['bewerb','hr','recruit','karrier','job','career'])]
    if pref: return pref[0]
    pref2 = [e for e in emails if any(k in e for k in ['info','kontakt','contact'])]
    if pref2: return pref2[0]
    valid = sorted(emails)
    return valid[0] if valid else ''

def try_impressum(base_url, safe_name):
    """Try impressum/kontakt for email."""
    for path in ['/impressum','/Impressum','/kontakt','/Kontakt','/contact']:
        url = base_url.rstrip('/') + path
        data, final_url, ct = url_get(url, timeout=8)
        if data and 'text/html' in (ct or ''):
            email = extract_emails(data)
            if email:
                fp = os.path.join(HTML_DIR, f"{safe_name}_imp.html")
                with open(fp, 'wb') as f: f.write(data)
                return email, url
    return '', ''

def main():
    start_idx = int(sys.argv[1]) if len(sys.argv) > 1 else 0
    batch_size = int(sys.argv[2]) if len(sys.argv) > 2 else 20
    
    log(f"=== z-ai Enrichment | start={start_idx} batch={batch_size} ===")
    
    with open(RESULTS_FILE) as f: all_results = json.load(f)
    filtered = [r for r in all_results if r.get('city') in TARGET_CITIES]
    
    enrichment = {}
    if os.path.exists(ENRICH_FILE):
        with open(ENRICH_FILE) as f: enrichment = json.load(f)
    
    need = [r for r in filtered if r.get('name','').strip() and r['name'].strip() not in enrichment]
    log(f"Need enrichment: {len(need)}")
    
    batch = need[start_idx:start_idx + batch_size]
    if not batch:
        log("Nothing to process!"); return
    
    ok = fail = no_web = emails_found = 0
    
    for i, firm in enumerate(batch):
        name = firm.get('name','').strip()
        city = firm.get('city','')
        
        try:
            log(f"[{i+1}/{len(batch)}] {name} ({city})")
            
            website = search_website_zai(name, city)
            
            if not website:
                no_web += 1
                enrichment[name] = {'website':'','html_file':'','email':'','status':'no_website','at':datetime.now().isoformat()[:19]}
                with open(ENRICH_FILE,'w') as f: json.dump(enrichment,f,ensure_ascii=False)
                time.sleep(1)
                continue
            
            safe = safe_filename(name)
            html_path = os.path.join(HTML_DIR, f"{safe}.html")
            rel_path = f"website_html/{safe}.html"
            
            data, final_url, ct = url_get(website, timeout=15)
            email = ''
            imp_url = ''
            html_size = 0
            
            if data and 'text/html' in (ct or ''):
                with open(html_path, 'wb') as f: f.write(data)
                html_size = len(data)
                email = extract_emails(data)
                
                if not email:
                    email, imp_url = try_impressum(final_url or website, safe)
                
                ok += 1
                if email: emails_found += 1
                log(f"  OK: {website[:60]} | {html_size}b | email={email or '-'}")
                enrichment[name] = {
                    'website': final_url or website,
                    'html_file': rel_path,
                    'email': email,
                    'impressum_url': imp_url,
                    'html_size': html_size,
                    'status': 'success',
                    'at': datetime.now().isoformat()[:19]
                }
            else:
                fail += 1
                enrichment[name] = {'website':website,'html_file':'','email':'','status':'html_fail','at':datetime.now().isoformat()[:19]}
            
            with open(ENRICH_FILE,'w') as f: json.dump(enrichment,f,ensure_ascii=False)
            time.sleep(1.5)
        
        except Exception as e:
            log(f"  ERROR: {str(e)[:80]}")
            enrichment[name] = {'website':'','html_file':'','email':'','status':'error','at':datetime.now().isoformat()[:19]}
            with open(ENRICH_FILE,'w') as f: json.dump(enrichment,f,ensure_ascii=False)
            time.sleep(2)
    
    log(f"Done! OK:{ok} Fail:{fail} NoWeb:{no_web} Emails:{emails_found} | Total:{len(enrichment)}")

if __name__ == '__main__':
    main()
