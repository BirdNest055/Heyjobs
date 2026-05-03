#!/usr/bin/env python3
"""
Lightweight Google Maps Scraper - one town at a time, proper cleanup
"""
import time, json, re, os, sys, gc, subprocess
from datetime import datetime
from playwright.sync_api import sync_playwright

RESULTS_FILE = '/home/z/my-project/gmaps_erlangen_results.json'
PROGRESS_FILE = '/home/z/my-project/gmaps_erlangen_progress.json'
LOG_FILE = '/home/z/my-project/gmaps_erlangen_log.txt'

ERLANGEN_LAT = 49.5969
ERLANGEN_LON = 11.0043

def log(msg):
    ts = datetime.now().strftime('%H:%M:%S')
    line = f"[{ts}] {msg}"
    print(line, flush=True)
    with open(LOG_FILE, 'a') as f:
        f.write(line + '\n')

FULL_TERMS = [
    "Firma", "Unternehmen", "GmbH", "IT Firma", "Software", "Softwareentwicklung",
    "Handwerk", "Industrie", "Handel", "Dienstleistung", "Gastronomie", "Baufirma",
    "Kfz Werkstatt", "Steuerberater", "Rechtsanwalt", "Versicherung", "Immobilien",
    "Marketing Agentur", "Logistik", "Elektro", "Metallbau", "Maschinenbau",
    "Kunststoff", "Druckerei", "Apotheke", "Arztpraxis", "Friseur", "Bäckerei",
]

MEDIUM_TERMS = [
    "Firma", "Unternehmen", "GmbH", "Handwerk", "Dienstleistung",
    "Gastronomie", "Baufirma", "IT", "Steuerberater", "Kfz", "Elektro",
    "Friseur", "Bäckerei",
]

MINIMAL_TERMS = ["Firma", "Unternehmen", "Gewerbe", "Handwerk", "Dienstleistung"]

def get_towns():
    with open('/home/z/my-project/gmaps_erlangen_towns.json', 'r') as f:
        return json.load(f)

def load_results():
    if os.path.exists(RESULTS_FILE):
        with open(RESULTS_FILE, 'r') as f:
            return json.load(f)
    return []

def save_results(results):
    with open(RESULTS_FILE, 'w') as f:
        json.dump(results, f, ensure_ascii=False)

def load_progress():
    if os.path.exists(PROGRESS_FILE):
        with open(PROGRESS_FILE, 'r') as f:
            return json.load(f)
    return {'completed_searches': []}

def save_progress(progress):
    with open(PROGRESS_FILE, 'w') as f:
        json.dump(progress, f, ensure_ascii=False)

def do_scrape(town_name, terms, all_results, seen_keys, progress):
    """Scrape one town with given search terms. Returns updated results."""
    
    # Kill old processes
    os.system('pkill -9 -f chromium 2>/dev/null')
    os.system('pkill -9 -f Xvfb 2>/dev/null')
    time.sleep(2)
    
    # Start fresh Xvfb
    os.system('Xvfb :99 -screen 0 1280x800x16 &')  # Less memory with 16-bit color
    time.sleep(2)
    os.environ['DISPLAY'] = ':99'
    
    p = sync_playwright().start()
    browser = None
    
    try:
        browser = p.chromium.launch(
            headless=False,
            args=[
                '--no-sandbox', 
                '--disable-blink-features=AutomationControlled',
                '--disable-dev-shm-usage', 
                '--disable-gpu',
                '--disable-software-rasterizer',
                '--js-flags=--max-old-space-size=256',
                '--disable-extensions',
                '--disable-notifications',
            ]
        )
        context = browser.new_context(
            viewport={'width': 1280, 'height': 800},
            user_agent='Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36',
            locale='de-DE'
        )
        context.add_init_script("""
            Object.defineProperty(navigator, 'webdriver', {get: () => undefined});
            window.chrome = {runtime: {}};
        """)
        page = context.new_page()
        
        total_new = 0
        
        for term in terms:
            search_key = f"{town_name}|{term}"
            if search_key in progress['completed_searches']:
                continue
            
            query = f'{term} in {town_name}'
            url = f'https://www.google.com/maps/search/{query.replace(" ", "+")}'
            
            try:
                page.goto(url, timeout=20000, wait_until='domcontentloaded')
            except Exception as e:
                log(f"  Load error: {str(e)[:60]}")
                continue
            
            time.sleep(2.5)
            
            # Handle cookies
            try:
                btn = page.query_selector('button[aria-label*="Alle ablehnen"]')
                if not btn:
                    btn = page.query_selector('button[aria-label*="Reject all"]')
                if btn:
                    btn.click()
                    time.sleep(0.8)
            except:
                pass
            
            # Check for CAPTCHA
            if "sorry" in page.url:
                log("  ⚠️ CAPTCHA! Stopping.")
                break
            
            # Scroll
            for _ in range(5):
                try:
                    page.evaluate('document.querySelector("div[role=feed]")?.scrollBy(0, 800)')
                    time.sleep(1.0)
                except:
                    break
            
            time.sleep(0.5)
            
            # Extract
            raw = page.evaluate("""() => {
                const businesses = [];
                const seen = new Set();
                document.querySelectorAll('a[href*="/maps/place/"]').forEach(link => {
                    try {
                        const href = link.getAttribute('href') || '';
                        const match = href.match(/\\/maps\\/place\\/([^/]+)/);
                        if (!match) return;
                        const slug = match[1];
                        if (seen.has(slug)) return;
                        seen.add(slug);
                        const ariaLabel = link.getAttribute('aria-label') || '';
                        let card = link;
                        for (let i = 0; i < 10; i++) {
                            if (card.parentElement) card = card.parentElement;
                            if ((card.innerText || '').length > 40) break;
                        }
                        businesses.push({
                            name: ariaLabel.substring(0, 150),
                            href: href,
                            text: (card.innerText || '').substring(0, 800),
                        });
                    } catch(e) {}
                });
                return businesses;
            }""") or []
            
            new_count = 0
            for r in raw:
                name = r.get('name', '').strip()
                key = name.lower().strip()
                if key and len(key) > 2 and key not in seen_keys:
                    text = r.get('text', '')
                    href = r.get('href', '')
                    
                    # Parse
                    rating = ''
                    m = re.search(r'(\d[,.]\d)', text[:50])
                    if m: rating = m.group(1).replace(',', '.')
                    
                    phone = ''
                    m = re.search(r'(\+49[\s/\-\d]{7,})', text)
                    if m: phone = m.group(1).strip()
                    
                    address = ''
                    m = re.search(r'([A-ZÄÖÜ][a-zäöüß]+(?:straße|str\.|str|weg|platz|allee|ring|gasse|damm|ufer|berg|feld|garten|hof|graben|markt)\s*\.?\s*\d+[a-zA-Z]?)', text)
                    if m: address = m.group(1).strip()
                    
                    has_website = bool(re.search(r'(Website|Webseite|Besuchen)', text, re.I))
                    
                    lat, lon = '', ''
                    m = re.search(r'!3d([\d.]+)!4d([\d.]+)', href)
                    if m: lat, lon = m.group(1), m.group(2)
                    
                    all_results.append({
                        'name': name, 'rating': rating, 'address': address,
                        'city': town_name, 'phone': phone, 'has_website': has_website,
                        'lat': lat, 'lon': lon, 'href': href,
                        'search_town': town_name, 'source': 'google_maps',
                        'scraped_at': datetime.now().isoformat()[:19],
                    })
                    seen_keys.add(key)
                    new_count += 1
            
            total_new += new_count
            progress['completed_searches'].append(search_key)
            log(f"  '{term}': {len(raw)} raw, {new_count} new | Total: {len(all_results)}")
            
            save_results(all_results)
            save_progress(progress)
            time.sleep(2)
    
    except Exception as e:
        log(f"  Error: {str(e)[:120]}")
    
    finally:
        try: context.close()
        except: pass
        try: browser.close()
        except: pass
        try: p.stop()
        except: pass
        os.system('pkill -9 -f chromium 2>/dev/null')
        os.system('pkill -9 -f Xvfb 2>/dev/null')
        gc.collect()
    
    return total_new

def main():
    start_town = int(sys.argv[1]) if len(sys.argv) > 1 else 0
    num_towns = int(sys.argv[2]) if len(sys.argv) > 2 else 5
    
    log(f"=== Erlangen 50km: Towns {start_town} to {start_town+num_towns-1} ===")
    
    towns = get_towns()
    all_results = load_results()
    progress = load_progress()
    seen_keys = set(r.get('name', '').lower().strip() for r in all_results if r.get('name'))
    
    log(f"Starting with {len(all_results)} existing results")
    
    end_town = min(start_town + num_towns, len(towns))
    
    for ti in range(start_town, end_town):
        town = towns[ti]
        town_name = town['name']
        tier = town['population_tier']
        dist = town['distance_km']
        
        terms = FULL_TERMS if tier <= 2 else (MEDIUM_TERMS if tier == 3 else MINIMAL_TERMS)
        
        # Count remaining searches for this town
        remaining = sum(1 for t in terms if f"{town_name}|{t}" not in progress['completed_searches'])
        if remaining == 0:
            log(f"Skip {town_name} - all searches completed")
            continue
        
        log(f"\n--- {ti+1}/{len(towns)}: {town_name} ({dist}km, tier={tier}, {remaining} remaining) ---")
        
        new = do_scrape(town_name, terms, all_results, seen_keys, progress)
        log(f"  Town done: {new} new businesses")
    
    log(f"\n=== Batch complete. Total: {len(all_results)} businesses ===")

if __name__ == '__main__':
    main()
