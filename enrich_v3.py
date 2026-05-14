#!/usr/bin/env python3
"""Website Enrichment V3 - uses z-ai web search API instead of DuckDuckGo."""
import json, re, time, os, sys, traceback, socket, subprocess, tempfile
from datetime import datetime
import urllib.request, urllib.parse, urllib.error

RESULTS_FILE = '/home/z/my-project/gmaps_erlangen_results.json'
ENRICH_FILE = '/home/z/my-project/website_enrichment.json'
HTML_DIR = '/home/z/my-project/download/website_html'
LOG_FILE = '/home/z/my-project/enrich_v3_log.txt'
TARGET_CITIES = ['Bamberg', 'Erlangen', 'Nürnberg']

# Domains to skip when looking for company websites
SKIP_DOMAINS = [
    'facebook.com','instagram.com','twitter.com','linkedin.com','youtube.com',
    'wikipedia.org','maps.google','gelbeseiten.de','yelp.de','tripadvisor',
    'firmenwissen.de','northdata','kompass','wlw.de','cylex.de','hotfrog',
    'cybo.com','dascleverle','kundeu.com','mojomox','opening-hours',
    'unternehmensregister.de','indofolio','bizdb','de.lusha','checkfacebook',
    'booking.com','11880.com','dasoertliche.de','check24.de','bloomberg.com',
    'crunchbase.com','glassdoor.com','kununu.com','indeed.com','stepstone.de',
    'meisterkarte.de','handwerksuche','diebestenderstadt.de','google.com',
    'provenexpert.com','dastelefonbuch.de','woobi.de','cylex-branchenbuch',
    'branchenbuch24.net','webwiki.de','mittelstandswiki.de','hub.6clicks.com',
    'jevee.de','bayerischewirtschaft.de','ihk-schwaben.de','ratemyarea.com',
]

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
                time.sleep(2)
            else:
                return None, None, None

def search_website(name, city):
    """Search for company website using z-ai web search API."""
    # Use single query for speed; fallback to quoted if needed
    queries = [f'{name} {city}']
    
    for qi, query in enumerate(queries):
        try:
            # Use z-ai CLI for web search
            tmp = tempfile.mktemp(suffix='.json')
            result = subprocess.run(
                ['z-ai', 'function', '-n', 'web_search', '-a', 
                 json.dumps({"query": query, "num": 5}), '-o', tmp],
                capture_output=True, text=True, timeout=15
            )
            
            if result.returncode != 0:
                log(f"  Search API error: {result.stderr[:80]}")
                continue
            
            with open(tmp) as f:
                results = json.load(f)
            
            # Clean up temp file
            try:
                os.unlink(tmp)
            except:
                pass
            
            if not results:
                continue
            
            # Filter results
            for item in results:
                url = item.get('url', '')
                host = item.get('host_name', '').lower()
                
                if not url.startswith('http'):
                    continue
                
                if any(s in url.lower() or s in host for s in SKIP_DOMAINS):
                    continue
                
                return url
            
            # If first query found no good results, try quoted query
            if qi == 0:
                queries.append(f'"{name}" {city} offizielle Webseite')
            
        except subprocess.TimeoutExpired:
            log(f"  Search timeout for: {query[:50]}")
        except Exception as e:
            log(f"  Search error: {str(e)[:60]}")
        
        time.sleep(0.5)
    
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
    batch_size = int(sys.argv[2]) if len(sys.argv) > 2 else 20
    
    log("=" * 50)
    log(f"Website Enrichment V3 | start={start_idx} batch={batch_size}")
    
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
            
            # Search using z-ai API
            website = search_website(name, city)
            
            if not website:
                no_web += 1
                enrichment[name] = {'website':'','html_file':'','email':'','impressum_url':'','impressum_file':'','status':'no_website','at':datetime.now().isoformat()[:19]}
                with open(ENRICH_FILE,'w') as f: json.dump(enrichment,f,ensure_ascii=False)
                time.sleep(0.5)
                continue
            
            # Download HTML
            safe = safe_filename(name)
            html_path = os.path.join(HTML_DIR, f"{safe}.html")
            rel_path = f"website_html/{safe}.html"
            
            data, final_url, ct = url_get(website, timeout=10, retries=1)
            
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
                    if email:
                        log(f"  Email from impressum: {email}")
                
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
            time.sleep(1)
        
        except Exception as e:
            log(f"  ERROR: {str(e)[:80]}")
            enrichment[name] = {'website':'','html_file':'','email':'','status':'error','error':str(e)[:100],'at':datetime.now().isoformat()[:19]}
            with open(ENRICH_FILE,'w') as f: json.dump(enrichment,f,ensure_ascii=False)
            time.sleep(1)
    
    log(f"\nDone! OK:{ok} Fail:{fail} NoWeb:{no_web} | Total enriched:{len(enrichment)}")

if __name__ == '__main__':
    main()
