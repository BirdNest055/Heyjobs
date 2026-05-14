#!/usr/bin/env python3
"""Website Enrichment V2 - more robust with retries and rate limiting."""
import json, re, time, os, sys, traceback, socket
from datetime import datetime
import urllib.request, urllib.parse, urllib.error

RESULTS_FILE = '/home/z/my-project/gmaps_erlangen_results.json'
ENRICH_FILE = '/home/z/my-project/website_enrichment.json'
HTML_DIR = '/home/z/my-project/download/website_html'
LOG_FILE = '/home/z/my-project/enrich_v2_log.txt'
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

def url_get(url, timeout=10, retries=2):
    """HTTP GET with retries."""
    for attempt in range(retries + 1):
        try:
            req = urllib.request.Request(url, headers={
                'User-Agent': 'Mozilla/5.0 (X11; Linux x86_64; rv:109.0) Gecko/20100101 Firefox/115.0',
                'Accept': 'text/html,application/xhtml+xml,*/*',
                'Accept-Language': 'de-DE,de;q=0.9,en;q=0.5',
            })
            with urllib.request.urlopen(req, timeout=timeout) as resp:
                return resp.read(), resp.url, resp.headers.get('Content-Type','')
        except (urllib.error.URLError, socket.timeout, ConnectionError, OSError) as e:
            if attempt < retries:
                time.sleep(3)
            else:
                return None, None, None

def search_website(name, city):
    """Search DuckDuckGo for company website."""
    queries = [
        f'{name} {city}',
        f'"{name}" {city} Webseite',
    ]
    
    for query in queries:
        url = f'https://html.duckduckgo.com/html/?q={urllib.parse.quote(query)}'
        data, _, _ = url_get(url, timeout=12, retries=1)
        if not data:
            continue
        
        html = data.decode('utf-8', errors='replace')
        links = re.findall(r'uddg=([^&"]+)', html)
        
        skip = ['facebook.com','instagram.com','twitter.com','linkedin.com','youtube.com',
                'wikipedia.org','maps.google','gelbeseiten.de','yelp.de','tripadvisor',
                'firmenwissen.de','northdata','kompass','wlw.de','cylex.de','hotfrog',
                'cybo.com','dascleverle','kundeu.com','mojomox','opening-hours',
                'unternehmensregister.de','indofolio','bizdb','de.lusha','checkfacebook']
        
        for link in links[:8]:
            decoded = urllib.parse.unquote(link)
            if not decoded.startswith('http'):
                continue
            if any(s in decoded.lower() for s in skip):
                continue
            return decoded
        
        time.sleep(1)
    
    return None

def extract_emails(html_bytes):
    """Extract emails from HTML."""
    try:
        text = html_bytes.decode('utf-8', errors='replace')
    except:
        text = html_bytes.decode('latin-1', errors='replace')
    
    emails = set()
    for m in re.finditer(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}', text):
        e = m.group(0).lower()
        if not any(x in e for x in ['.png','.jpg','.gif','.svg','.css','.js','example.com','email.com','domain.com','sentry','wixpress','googlemail']):
            emails.add(e)
    
    for m in re.finditer(r'mailto:([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})', text, re.I):
        emails.add(m.group(1).lower())
    
    for m in re.finditer(r'[a-zA-Z0-9._%+-]+\s*[\(\[]at[\)\]]\s*[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}', text, re.I):
        emails.add(re.sub(r'\s*[\(\[]at[\)\]]\s*', '@', m.group(0).lower()))
    
    skip_emails = ['sage','wordpress','admin@localhost','test@','noreply@','example@','webmaster@','postmaster@','donotreply@']
    emails = {e for e in emails if not any(s in e for s in skip_emails)}
    
    # Prefer bewerbung/hr emails
    pref = [e for e in emails if any(k in e for k in ['bewerb','hr','recruit','karrier','job','career'])]
    if pref: return pref[0]
    pref2 = [e for e in emails if any(k in e for k in ['info','kontakt','contact'])]
    if pref2: return pref2[0]
    valid = sorted(emails)
    return valid[0] if valid else ''

def try_impressum(base_url, safe_name):
    """Try impressum/kontakt pages for email."""
    paths = ['/impressum','/Impressum','/kontakt','/Kontakt','/contact','/impressum.html','/kontakt.html']
    base = base_url.rstrip('/')
    
    for path in paths:
        url = base + path
        data, final_url, ct = url_get(url, timeout=8, retries=1)
        if data and 'text/html' in (ct or ''):
            email = extract_emails(data)
            if email:
                fp = os.path.join(HTML_DIR, f"{safe_name}_imp{path.replace('/','_')}.html")
                with open(fp, 'wb') as f:
                    f.write(data)
                return email, url, f"website_html/{safe_name}_imp{path.replace('/','_')}.html"
    return '', '', ''

def main():
    start_idx = int(sys.argv[1]) if len(sys.argv) > 1 else 0
    batch_size = int(sys.argv[2]) if len(sys.argv) > 2 else 50
    
    log("=" * 50)
    log(f"Website Enrichment V2 | start={start_idx} batch={batch_size}")
    
    with open(RESULTS_FILE) as f:
        all_results = json.load(f)
    
    filtered = [r for r in all_results if r.get('city') in TARGET_CITIES]
    
    enrichment = {}
    if os.path.exists(ENRICH_FILE):
        with open(ENRICH_FILE) as f:
            enrichment = json.load(f)
    
    need = [r for r in filtered if r.get('name','').strip() and r['name'].strip() not in enrichment]
    need.sort(key=lambda x: x.get('city',''))
    log(f"Need enrichment: {len(need)}")
    
    batch = need[start_idx:start_idx + batch_size]
    if not batch:
        log("Nothing to process!")
        return
    
    ok = fail = no_web = 0
    
    for i, firm in enumerate(batch):
        name = firm.get('name','').strip()
        city = firm.get('city','')
        
        try:
            log(f"[{i+1}/{len(batch)}] {name} ({city})")
            
            # Search
            website = search_website(name, city)
            
            if not website:
                no_web += 1
                enrichment[name] = {'website':'','html_file':'','email':'','impressum_url':'','impressum_file':'','status':'no_website','at':datetime.now().isoformat()[:19]}
                with open(ENRICH_FILE,'w') as f: json.dump(enrichment,f,ensure_ascii=False)
                time.sleep(2)
                continue
            
            # Download HTML
            safe = safe_filename(name)
            html_path = os.path.join(HTML_DIR, f"{safe}.html")
            rel_path = f"website_html/{safe}.html"
            
            data, final_url, ct = url_get(website, timeout=15, retries=2)
            
            email = ''
            imp_url = ''
            imp_file = ''
            html_size = 0
            
            if data and 'text/html' in (ct or ''):
                with open(html_path, 'wb') as f:
                    f.write(data)
                html_size = len(data)
                
                email = extract_emails(data)
                
                if not email:
                    email, imp_url, imp_file = try_impressum(final_url or website, safe)
                
                ok += 1
                log(f"  OK: {website} | {html_size}b | email={email or '-'}")
                enrichment[name] = {
                    'website': final_url or website,
                    'html_file': rel_path,
                    'email': email,
                    'impressum_url': imp_url,
                    'impressum_file': imp_file,
                    'html_size': html_size,
                    'status': 'success',
                    'at': datetime.now().isoformat()[:19]
                }
            else:
                fail += 1
                enrichment[name] = {'website':website,'html_file':'','email':'','status':'html_fail','at':datetime.now().isoformat()[:19]}
            
            with open(ENRICH_FILE,'w') as f: json.dump(enrichment,f,ensure_ascii=False)
            time.sleep(2)
        
        except Exception as e:
            log(f"  ERROR: {str(e)[:80]}")
            enrichment[name] = {'website':'','html_file':'','email':'','status':'error','error':str(e)[:100],'at':datetime.now().isoformat()[:19]}
            with open(ENRICH_FILE,'w') as f: json.dump(enrichment,f,ensure_ascii=False)
            time.sleep(3)
    
    log(f"\nDone! OK:{ok} Fail:{fail} NoWeb:{no_web} | Total enriched:{len(enrichment)}")

if __name__ == '__main__':
    main()
