#!/usr/bin/env python3
"""Final Google Maps scraper - processes remaining towns in a single Playwright session."""
import time, json, re, os, sys, traceback
from datetime import datetime
os.environ["DISPLAY"] = ":98"
from playwright.sync_api import sync_playwright

RESULTS_FILE = '/home/z/my-project/gmaps_erlangen_results.json'
PROGRESS_FILE = '/home/z/my-project/gmaps_erlangen_progress.json'

MINIMAL = ["Firma", "Unternehmen", "Gewerbe", "Handwerk", "Dienstleistung"]
MEDIUM = MINIMAL + ["GmbH", "Gastronomie", "Baufirma", "IT", "Steuerberater", "Kfz", "Elektro", "Friseur", "Bäckerei"]
FULL = MEDIUM + ["Software", "Softwareentwicklung", "Industrie", "Handel", "Rechtsanwalt", "Versicherung", "Immobilien", "Marketing Agentur", "Logistik", "Metallbau", "Maschinenbau", "Kunststoff", "Druckerei", "Apotheke", "Arztpraxis"]

def log(msg):
    print(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}", flush=True)

def main():
    with open('/home/z/my-project/gmaps_erlangen_towns.json') as f: all_towns = json.load(f)
    with open(RESULTS_FILE) as f: all_results = json.load(f)
    with open(PROGRESS_FILE) as f: progress = json.load(f)
    
    done_towns = set(cs.split('|')[0] for cs in progress['completed_searches'])
    remaining = [t for t in all_towns if t['name'] not in done_towns]
    remaining.sort(key=lambda x: x['distance_km'])
    
    log(f"Remaining: {len(remaining)} towns, Existing: {len(all_results)} results")
    if not remaining:
        log("All done!")
        return
    
    seen = set(r.get('name','').lower().strip() for r in all_results if r.get('name'))
    
    p = sync_playwright().start()
    browser = p.chromium.launch(headless=False, args=['--no-sandbox', '--disable-blink-features=AutomationControlled', '--disable-dev-shm-usage', '--disable-gpu'])
    ctx = browser.new_context(viewport={'width':1920,'height':1080}, user_agent='Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36', locale='de-DE')
    ctx.add_init_script('Object.defineProperty(navigator, "webdriver", {get: () => undefined});')
    page = ctx.new_page()
    
    total_new = 0
    errors = 0
    
    try:
        for town in remaining:
            name = town['name']
            tier = town['population_tier']
            terms = FULL if tier <= 2 else (MEDIUM if tier == 3 else MINIMAL)
            log(f"\n=== {name} ({town['distance_km']}km, tier={tier}, {len(terms)} terms) ===")
            
            for term in terms:
                sk = f"{name}|{term}"
                if sk in progress['completed_searches']:
                    continue
                
                q = f'{term} in {name}'
                log(f"  {q}")
                
                try:
                    page.goto(f'https://www.google.com/maps/search/{q.replace(" ", "+")}', timeout=25000, wait_until='domcontentloaded')
                    time.sleep(3)
                    
                    # Cookies
                    for _ in range(3):
                        btn = page.query_selector('button[aria-label*="Alle ablehnen"]') or page.query_selector('button[aria-label*="Reject all"]')
                        if btn:
                            try: btn.click(); time.sleep(1)
                            except: pass
                            break
                        time.sleep(0.5)
                    
                    if "sorry" in page.url:
                        log("  CAPTCHA! Wait 60s..."); time.sleep(60); continue
                    
                    for _ in range(6):
                        try:
                            page.evaluate('document.querySelector("div[role=feed]")?.scrollBy(0, 1000)')
                            time.sleep(1.5)
                        except: break
                    time.sleep(1)
                    
                    raw = page.evaluate(r"""() => {
                        const b = [], s = new Set();
                        document.querySelectorAll('a[href*="/maps/place/"]').forEach(l => {
                            try {
                                const h = l.getAttribute('href') || '';
                                const m = h.match(/\/maps\/place\/([^/]+)/);
                                if (!m || s.has(m[1])) return;
                                s.add(m[1]);
                                const a = l.getAttribute('aria-label') || '';
                                let c = l;
                                for (let i = 0; i < 10; i++) { if (c.parentElement) c = c.parentElement; if ((c.innerText||'').length > 40) break; }
                                const t = c.innerText || '';
                                b.push({name:a.substring(0,150), href:h, text:t.substring(0,800)});
                            } catch(e) {}
                        });
                        return b;
                    }""")
                    
                    new = 0
                    for r in (raw or []):
                        n = r.get('name','').strip()
                        k = n.lower()
                        if k and len(k) > 2 and k not in seen:
                            text = r.get('text','')
                            href = r.get('href','')
                            
                            rating = ''
                            m = re.search(r'(\d[,.]\d)', text[:50])
                            if m: rating = m.group(1).replace(',', '.')
                            
                            phone = ''
                            m = re.search(r'(\+49[\s/\-\d]{7,}|\d{3,4}[\s/\-]\d{5,8})', text)
                            if m: phone = m.group(1).strip()
                            
                            address = ''
                            m = re.search(r'([A-ZÄÖÜ][a-zäöüß]+(?:straße|str\.|str|weg|platz|allee|ring|gasse|damm|ufer|berg|feld|garten|hof|graben|steg|markt|kirch|park|see|tal|bach)\s*\.?\s*\d+[a-zA-Z]?)', text)
                            if m: address = m.group(1).strip()
                            
                            plz = ''
                            m = re.search(r'\b(\d{5})\s+[A-ZÄÖÜ]', text)
                            if m: plz = m.group(1)
                            if not plz:
                                m = re.search(r'\b(9[0-5]\d{3})\b', text)
                                if m: plz = m.group(1)
                            
                            city = name
                            m = re.search(r'\b\d{5}\s+([A-ZÄÖÜ][a-zäöüß]+(?:\s+[A-ZÄÖÜ][a-zäöüß]+)?)', text)
                            if m and len(m.group(1).strip()) > 2: city = m.group(1).strip()
                            
                            has_website = bool(re.search(r'(Website|Webseite|website|Besuchen|Visit)', text, re.I))
                            
                            lat, lon = '', ''
                            m = re.search(r'!3d([\d.]+)!4d([\d.]+)', href)
                            if m: lat, lon = m.group(1), m.group(2)
                            
                            all_results.append({
                                'name':n, 'rating':rating, 'address':address, 'plz':plz, 'city':city,
                                'phone':phone, 'has_website':has_website, 'lat':lat, 'lon':lon,
                                'href':href, 'search_town':name, 'source':'google_maps',
                                'scraped_at': datetime.now().isoformat()[:19],
                            })
                            seen.add(k)
                            new += 1
                    
                    total_new += new
                    errors = 0
                    log(f"  {len(raw or [])} listings, {new} new | Total: {len(all_results)}")
                    
                    progress['completed_searches'].append(sk)
                    with open(RESULTS_FILE, 'w') as f: json.dump(all_results, f, ensure_ascii=False, indent=1)
                    with open(PROGRESS_FILE, 'w') as f: json.dump(progress, f, ensure_ascii=False)
                    
                    time.sleep(3)
                
                except Exception as e:
                    errors += 1
                    log(f"  Error: {str(e)[:80]}")
                    if errors >= 10:
                        log("Too many errors, stopping!"); break
                    # Try browser restart
                    try: page.evaluate('1+1')
                    except:
                        log("  Restarting browser...")
                        try: browser.close()
                        except: pass
                        try: p.stop()
                        except: pass
                        time.sleep(3)
                        p = sync_playwright().start()
                        browser = p.chromium.launch(headless=False, args=['--no-sandbox','--disable-blink-features=AutomationControlled','--disable-dev-shm-usage','--disable-gpu'])
                        ctx = browser.new_context(viewport={'width':1920,'height':1080}, user_agent='Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36', locale='de-DE')
                        ctx.add_init_script('Object.defineProperty(navigator, "webdriver", {get: () => undefined});')
                        page = ctx.new_page()
                    time.sleep(5)
            
            log(f"  Town done. Total: {len(all_results)}")
            if errors >= 10: break
    
    except Exception as e:
        log(f"Fatal: {str(e)[:120]}")
        traceback.print_exc()
    finally:
        with open(RESULTS_FILE, 'w') as f: json.dump(all_results, f, ensure_ascii=False, indent=1)
        with open(PROGRESS_FILE, 'w') as f: json.dump(progress, f, ensure_ascii=False)
        try: browser.close()
        except: pass
        try: p.stop()
        except: pass
    
    log(f"\nDone! +{total_new} new, Total: {len(all_results)}")

if __name__ == '__main__':
    main()
