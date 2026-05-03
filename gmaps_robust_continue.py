#!/usr/bin/env python3
"""
Robust Google Maps scraper - processes one town at a time, saves after each.
Auto-restarts on crashes. Designed to be called repeatedly.
"""
import time, json, re, os, sys, traceback
from datetime import datetime
from playwright.sync_api import sync_playwright

RESULTS_FILE = '/home/z/my-project/gmaps_erlangen_results.json'
PROGRESS_FILE = '/home/z/my-project/gmaps_erlangen_progress.json'

ERLANGEN_LAT = 49.5969
ERLANGEN_LON = 11.0043

MINIMAL_TERMS = ["Firma", "Unternehmen", "Gewerbe", "Handwerk", "Dienstleistung"]
MEDIUM_TERMS = MINIMAL_TERMS + ["GmbH", "Gastronomie", "Baufirma", "IT", "Steuerberater", "Kfz", "Elektro", "Friseur", "Bäckerei"]
FULL_TERMS = MEDIUM_TERMS + ["Software", "Softwareentwicklung", "Industrie", "Handel", "Rechtsanwalt", "Versicherung", "Immobilien", "Marketing Agentur", "Logistik", "Metallbau", "Maschinenbau", "Kunststoff", "Druckerei", "Apotheke", "Arztpraxis"]

def log(msg):
    ts = datetime.now().strftime('%H:%M:%S')
    print(f"[{ts}] {msg}", flush=True)

def search_and_extract(page, query, max_scrolls=8):
    """Search Google Maps and extract listings."""
    url = f'https://www.google.com/maps/search/{query.replace(" ", "+")}'
    try:
        page.goto(url, timeout=25000, wait_until='domcontentloaded')
    except Exception as e:
        log(f"  Load error: {str(e)[:60]}")
        return []
    
    time.sleep(3)
    
    # Handle cookies
    for _ in range(3):
        btn = page.query_selector('button[aria-label*="Alle ablehnen"]') or page.query_selector('button[aria-label*="Reject all"]')
        if btn:
            try:
                btn.click()
                time.sleep(1)
            except: pass
            break
        time.sleep(0.5)
    
    # Check CAPTCHA
    if "sorry" in page.url or "captcha" in page.url.lower():
        log("  CAPTCHA detected! Waiting 60s...")
        time.sleep(60)
        return []
    
    # Scroll
    for _ in range(max_scrolls):
        try:
            page.evaluate('document.querySelector("div[role=feed]")?.scrollBy(0, 1000)')
            time.sleep(1.5)
        except:
            break
    
    time.sleep(1)
    
    # Extract
    results = page.evaluate(r"""() => {
        const businesses = [];
        const seen = new Set();
        document.querySelectorAll('a[href*="/maps/place/"]').forEach(link => {
            try {
                const href = link.getAttribute('href') || '';
                const match = href.match(/\/maps\/place\/([^/]+)/);
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
                const text = card.innerText || '';
                const lines = text.split('\n').filter(l => l.trim());
                businesses.push({
                    name: ariaLabel.substring(0, 150),
                    href: href,
                    text: text.substring(0, 800),
                    lines: lines.slice(0, 12),
                });
            } catch(e) {}
        });
        return businesses;
    }""")
    
    return results or []

def parse_business(raw, town_name):
    text = raw.get('text', '')
    lines = raw.get('lines', [])
    name = raw.get('name', '')
    href = raw.get('href', '')
    
    rating = ''
    m = re.search(r'(\d[,.]\d)', text[:50])
    if m: rating = m.group(1).replace(',', '.')
    
    phone = ''
    m = re.search(r'(\+49[\s/\-\d]{7,}|\d{3,4}[\s/\-]\d{5,8})', text)
    if m: phone = m.group(1).strip()
    
    address = ''
    m = re.search(r'([A-ZÄÖÜ][a-zäöüß]+(?:straße|str\.|str|weg|platz|allee|ring|gasse|damm|ufer|berg|feld|garten|hof|graben|steg|markt|kirch|schloss|park|see|tal|bach)\s*\.?\s*\d+[a-zA-Z]?(?:\s*[-/]\s*\d+[a-zA-Z]?)?)', text)
    if m: address = m.group(1).strip()
    
    plz = ''
    m = re.search(r'\b(\d{5})\s+[A-ZÄÖÜ]', text)
    if m: plz = m.group(1)
    if not plz:
        m = re.search(r'\b(9[0-5]\d{3})\b', text)
        if m: plz = m.group(1)
    
    city = town_name
    m = re.search(r'\b\d{5}\s+([A-ZÄÖÜ][a-zäöüß]+(?:\s+[A-ZÄÖÜ][a-zäöüß]+)?)', text)
    if m and len(m.group(1).strip()) > 2: city = m.group(1).strip()
    
    has_website = bool(re.search(r'(Website|Webseite|website|Besuchen|Visit)', text, re.I))
    
    lat, lon = '', ''
    m = re.search(r'!3d([\d.]+)!4d([\d.]+)', href)
    if m: lat, lon = m.group(1), m.group(2)
    
    return {
        'name': name.strip(), 'rating': rating, 'address': address,
        'plz': plz, 'city': city, 'phone': phone, 'has_website': has_website,
        'lat': lat, 'lon': lon, 'href': href, 'search_town': town_name,
        'source': 'google_maps', 'scraped_at': datetime.now().isoformat()[:19],
    }

def main():
    max_towns = int(sys.argv[1]) if len(sys.argv) > 1 else 5
    
    # Load data
    with open('/home/z/my-project/gmaps_erlangen_towns.json') as f:
        all_towns = json.load(f)
    with open(RESULTS_FILE) as f:
        all_results = json.load(f)
    with open(PROGRESS_FILE) as f:
        progress = json.load(f)
    
    # Find remaining
    completed_town_names = set(cs.split('|')[0] for cs in progress['completed_searches'])
    remaining = [t for t in all_towns if t['name'] not in completed_town_names]
    remaining.sort(key=lambda x: x['distance_km'])
    
    log(f"Remaining: {len(remaining)} towns, Existing: {len(all_results)} results")
    
    if not remaining:
        log("All done!")
        return
    
    # Limit
    to_process = remaining[:max_towns]
    log(f"Processing {len(to_process)} towns this run")
    
    # Dedup
    seen_keys = set(r.get('name', '').lower().strip() for r in all_results if r.get('name'))
    
    # Start browser
    os.system('pkill Xvfb 2>/dev/null')
    time.sleep(1)
    os.system('Xvfb :99 -screen 0 1920x1080x24 &')
    time.sleep(2)
    os.environ['DISPLAY'] = ':99'
    
    p = sync_playwright().start()
    browser = p.chromium.launch(
        headless=False,
        args=['--no-sandbox', '--disable-blink-features=AutomationControlled', '--disable-dev-shm-usage']
    )
    ctx = browser.new_context(
        viewport={'width': 1920, 'height': 1080},
        user_agent='Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36',
        locale='de-DE'
    )
    ctx.add_init_script('Object.defineProperty(navigator, "webdriver", {get: () => undefined});')
    page = ctx.new_page()
    
    total_new = 0
    
    try:
        for town in to_process:
            town_name = town['name']
            tier = town['population_tier']
            
            terms = FULL_TERMS if tier <= 2 else (MEDIUM_TERMS if tier == 3 else MINIMAL_TERMS)
            
            log(f"\n=== {town_name} ({town['distance_km']}km, tier={tier}, {len(terms)} terms) ===")
            
            for term in terms:
                search_key = f"{town_name}|{term}"
                if search_key in progress['completed_searches']:
                    continue
                
                query = f'{term} in {town_name}'
                log(f"  Searching: '{query}'")
                
                try:
                    raw = search_and_extract(page, query)
                    new = 0
                    for r in raw:
                        parsed = parse_business(r, town_name)
                        key = parsed['name'].lower().strip()
                        if key and len(key) > 2 and key not in seen_keys:
                            all_results.append(parsed)
                            seen_keys.add(key)
                            new += 1
                    
                    total_new += new
                    log(f"  {len(raw)} listings, {new} new | Total: {len(all_results)}")
                    
                    progress['completed_searches'].append(search_key)
                    
                    # Save every search
                    with open(RESULTS_FILE, 'w') as f:
                        json.dump(all_results, f, ensure_ascii=False, indent=1)
                    with open(PROGRESS_FILE, 'w') as f:
                        json.dump(progress, f, ensure_ascii=False)
                    
                    time.sleep(3)
                except Exception as e:
                    log(f"  Error: {str(e)[:80]}")
                    try:
                        page.evaluate('1+1')
                    except:
                        log("  Browser dead, restarting...")
                        try: browser.close()
                        except: pass
                        try: p.stop()
                        except: pass
                        time.sleep(3)
                        p = sync_playwright().start()
                        browser = p.chromium.launch(headless=False, args=['--no-sandbox', '--disable-blink-features=AutomationControlled', '--disable-dev-shm-usage'])
                        ctx = browser.new_context(viewport={'width': 1920, 'height': 1080}, user_agent='Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36', locale='de-DE')
                        ctx.add_init_script('Object.defineProperty(navigator, "webdriver", {get: () => undefined});')
                        page = ctx.new_page()
                    time.sleep(5)
            
            log(f"  Town done. Total: {len(all_results)}")
    
    except Exception as e:
        log(f"Fatal: {str(e)[:120]}")
        traceback.print_exc()
    finally:
        try: browser.close()
        except: pass
        try: p.stop()
        except: pass
        os.system('pkill Xvfb 2>/dev/null')
    
    with open(RESULTS_FILE, 'w') as f:
        json.dump(all_results, f, ensure_ascii=False, indent=1)
    with open(PROGRESS_FILE, 'w') as f:
        json.dump(progress, f, ensure_ascii=False)
    
    log(f"\nSession done. +{total_new} new, Total: {len(all_results)}")

if __name__ == '__main__':
    main()
