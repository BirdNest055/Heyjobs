#!/usr/bin/env python3
"""Website Enrichment V4 - fast version: search first, download second, skip impressum."""
import json, re, time, os, sys, subprocess, tempfile
from datetime import datetime
import urllib.request, urllib.parse, urllib.error, socket

RESULTS_FILE = '/home/z/my-project/gmaps_erlangen_results.json'
ENRICH_FILE = '/home/z/my-project/website_enrichment.json'
HTML_DIR = '/home/z/my-project/download/website_html'
LOG_FILE = '/home/z/my-project/enrich_v4_log.txt'
TARGET_CITIES = ['Bamberg', 'Erlangen', 'Nürnberg']

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
    'branchenbuch24.net','webwiki.de','mittelstandswiki.de','jevee.de',
    'bayerischewirtschaft.de','ratemyarea.com','tellows.de','compuware.com',
    'wlb.de','lokal.blue','meinestadt.de','frag-die-ihk.de','ihk-nuernberg.de',
    'hellowork.com','stellenonline.de','jobware.de','arbeitsagentur.de',
    'azubiyo.de','ausbildung.de','bixploit.com','gebaeudereiniger-portal.de',
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

def search_website(name, city):
    """Search for company website using z-ai web search API."""
    query = f'{name} {city}'
    try:
        tmp = tempfile.mktemp(suffix='.json')
        result = subprocess.run(
            ['z-ai', 'function', '-n', 'web_search', '-a', 
             json.dumps({"query": query, "num": 5}), '-o', tmp],
            capture_output=True, text=True, timeout=15
        )
        
        if result.returncode != 0:
            return None
        
        with open(tmp) as f:
            results = json.load(f)
        
        try:
            os.unlink(tmp)
        except:
            pass
        
        if not results:
            return None
        
        for item in results:
            url = item.get('url', '')
            host = item.get('host_name', '').lower()
            if not url.startswith('http'):
                continue
            if any(s in url.lower() or s in host for s in SKIP_DOMAINS):
                continue
            return url
        
        # If all results are filtered, try quoted query
        query2 = f'"{name}" {city} website'
        tmp = tempfile.mktemp(suffix='.json')
        result = subprocess.run(
            ['z-ai', 'function', '-n', 'web_search', '-a',
             json.dumps({"query": query2, "num": 3}), '-o', tmp],
            capture_output=True, text=True, timeout=15
        )
        if result.returncode == 0:
            with open(tmp) as f:
                results2 = json.load(f)
            try:
                os.unlink(tmp)
            except:
                pass
            if results2:
                for item in results2:
                    url = item.get('url', '')
                    host = item.get('host_name', '').lower()
                    if not url.startswith('http'):
                        continue
                    if any(s in url.lower() or s in host for s in SKIP_DOMAINS):
                        continue
                    return url
        return None
        
    except subprocess.TimeoutExpired:
        return None
    except Exception as e:
        return None

def download_and_extract(website, safe_name):
    """Download HTML, save, and extract email. Fast version."""
    try:
        req = urllib.request.Request(website, headers={
            'User-Agent': 'Mozilla/5.0 (X11; Linux x86_64; rv:109.0) Gecko/20100101 Firefox/115.0',
            'Accept': 'text/html,application/xhtml+xml,*/*',
            'Accept-Language': 'de-DE,de;q=0.9,en;q=0.5',
        })
        with urllib.request.urlopen(req, timeout=10) as resp:
            ct = resp.headers.get('Content-Type','')
            if 'text/html' not in ct:
                return None, '', 0
            data = resp.read(500000)  # Max 500KB
            final_url = resp.url
    except Exception:
        return None, '', 0
    
    # Save HTML
    html_path = os.path.join(HTML_DIR, f"{safe_name}.html")
    with open(html_path, 'wb') as f:
        f.write(data)
    
    # Extract email
    try:
        text = data.decode('utf-8', errors='replace')
    except:
        text = data.decode('latin-1', errors='replace')
    
    emails = set()
    for m in re.finditer(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}', text):
        e = m.group(0).lower()
        if not any(x in e for x in ['.png','.jpg','.gif','.svg','.css','.js','example.com','email.com','domain.com','sentry','wixpress','googlemail']):
            emails.add(e)
    
    for m in re.finditer(r'mailto:([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})', text, re.I):
        emails.add(m.group(1).lower())
    
    skip_emails = ['sage','wordpress','admin@localhost','test@','noreply@','example@','webmaster@','postmaster@','donotreply@']
    emails = {e for e in emails if not any(s in e for s in skip_emails)}
    
    # Prefer bewerbung/hr/info emails
    pref = [e for e in emails if any(k in e for k in ['bewerb','hr','recruit','karrier','job','career'])]
    if pref: return final_url, pref[0], len(data)
    pref2 = [e for e in emails if any(k in e for k in ['info','kontakt','contact'])]
    if pref2: return final_url, pref2[0], len(data)
    valid = sorted(emails)
    email = valid[0] if valid else ''
    
    # If no email on main page, try impressum quickly (just 2 paths)
    if not email:
        for path in ['/impressum', '/kontakt']:
            try:
                imp_url = final_url.rstrip('/') + path
                req2 = urllib.request.Request(imp_url, headers={
                    'User-Agent': 'Mozilla/5.0 (X11; Linux x86_64; rv:109.0) Gecko/20100101 Firefox/115.0',
                })
                with urllib.request.urlopen(req2, timeout=6) as resp2:
                    if 'text/html' in resp2.headers.get('Content-Type',''):
                        imp_data = resp2.read(300000)
                        imp_text = imp_data.decode('utf-8', errors='replace')
                        imp_emails = set()
                        for m in re.finditer(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}', imp_text):
                            em = m.group(0).lower()
                            if not any(x in em for x in ['.png','.jpg','.gif','.svg','.css','.js','example.com','sentry','wixpress']):
                                imp_emails.add(em)
                        imp_emails = {e for e in imp_emails if not any(s in e for s in skip_emails)}
                        if imp_emails:
                            pref2 = [e for e in imp_emails if any(k in e for k in ['info','kontakt','contact'])]
                            email = pref2[0] if pref2 else sorted(imp_emails)[0]
                            # Save impressum page too
                            imp_path = os.path.join(HTML_DIR, f"{safe_name}_imp{path.replace('/','_')}.html")
                            with open(imp_path, 'wb') as f:
                                f.write(imp_data)
                            break
            except Exception:
                continue
    
    return final_url, email, len(data)

def main():
    start_idx = int(sys.argv[1]) if len(sys.argv) > 1 else 0
    batch_size = int(sys.argv[2]) if len(sys.argv) > 2 else 25
    
    log("=" * 50)
    log(f"Website Enrichment V4 | start={start_idx} batch={batch_size}")
    
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
        safe = safe_filename(name)
        
        try:
            log(f"[{i+1}/{len(batch)}] {name} ({city})")
            
            # Step 1: Search for website
            website = search_website(name, city)
            
            if not website:
                no_web += 1
                enrichment[name] = {
                    'website':'','html_file':'','email':'',
                    'status':'no_website',
                    'at':datetime.now().isoformat()[:19]
                }
                with open(ENRICH_FILE,'w') as f: json.dump(enrichment,f,ensure_ascii=False)
                continue
            
            # Step 2: Download and extract
            final_url, email, html_size = download_and_extract(website, safe)
            
            if final_url:
                ok += 1
                log(f"  OK: {final_url[:60]} | {html_size}b | email={email or '-'}")
                enrichment[name] = {
                    'website': final_url,
                    'html_file': f"website_html/{safe}.html",
                    'email': email,
                    'html_size': html_size,
                    'status': 'success',
                    'at': datetime.now().isoformat()[:19]
                }
            else:
                fail += 1
                enrichment[name] = {
                    'website': website, 'html_file':'', 'email':'',
                    'status':'html_fail',
                    'at':datetime.now().isoformat()[:19]
                }
            
            with open(ENRICH_FILE,'w') as f: json.dump(enrichment,f,ensure_ascii=False)
        
        except Exception as e:
            log(f"  ERROR: {str(e)[:80]}")
            enrichment[name] = {
                'website':'','html_file':'','email':'',
                'status':'error','error':str(e)[:100],
                'at':datetime.now().isoformat()[:19]
            }
            with open(ENRICH_FILE,'w') as f: json.dump(enrichment,f,ensure_ascii=False)
    
    log(f"\nDone! OK:{ok} Fail:{fail} NoWeb:{no_web} | Total enriched:{len(enrichment)}")

if __name__ == '__main__':
    main()
