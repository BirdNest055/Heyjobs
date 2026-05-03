#!/usr/bin/env python3
import os, sys, signal, traceback, time

# Signal handler
def handler(signum, frame):
    with open('/home/z/my-project/gmaps_signal.txt', 'a') as f:
        f.write(f"Signal {signum} received at {time.strftime('%H:%M:%S')}\n")
    sys.exit(1)

signal.signal(signal.SIGTERM, handler)
signal.signal(signal.SIGINT, handler)
signal.signal(signal.SIGHUP, handler)

try:
    from playwright.sync_api import sync_playwright
    
    print("Starting headless browser...", flush=True)
    p = sync_playwright().start()
    browser = p.chromium.launch(headless=True, args=['--no-sandbox','--disable-dev-shm-usage'])
    ctx = browser.new_context(viewport={'width':1920,'height':1080}, locale='de-DE')
    page = ctx.new_page()
    
    print("Navigating...", flush=True)
    page.goto('https://www.google.com/maps/search/Firma+in+Willhermsdorf', timeout=20000)
    time.sleep(3)
    print(f"Title: {page.title()}", flush=True)
    
    # Scroll
    for i in range(5):
        try:
            page.evaluate('document.querySelector("div[role=feed]")?.scrollBy(0, 1000)')
            time.sleep(1.5)
        except:
            break
    
    # Extract
    results = page.evaluate(r"""() => {
        const b = [], s = new Set();
        document.querySelectorAll('a[href*="/maps/place/"]').forEach(l => {
            try {
                const h = l.getAttribute('href') || '';
                const m = h.match(/\/maps\/place\/([^/]+)/);
                if (!m || s.has(m[1])) return;
                s.add(m[1]);
                const a = l.getAttribute('aria-label') || '';
                b.push(a.substring(0, 80));
            } catch(e) {}
        });
        return b;
    }""")
    
    print(f"Found {len(results)} listings", flush=True)
    for r in results[:5]:
        print(f"  - {r}", flush=True)
    
    browser.close()
    p.stop()
    print("SUCCESS!", flush=True)

except Exception as e:
    print(f"ERROR: {e}", flush=True)
    traceback.print_exc()
