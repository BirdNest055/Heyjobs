#!/usr/bin/env python3
"""Phase 2: Enrich employers with contact details from detail API"""
import json, time, random, requests, sys
sys.stdout.reconfigure(line_buffering=True)

DETAIL_API = 'https://rest.arbeitsagentur.de/jobboerse/jobsuche-service/pc/v2/jobdetails'
HEADERS = {
    'Accept': 'application/json, text/plain, */*',
    'X-API-Key': 'jobboerse-jobsuche',
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
    'Origin': 'https://www.arbeitsagentur.de', 'Referer': 'https://www.arbeitsagentur.de/',
}
OUTPUT = '/home/z/my-project/download/arbeitsagentur_employers.json'

with open(OUTPUT, 'r', encoding='utf-8') as f:
    data = json.load(f)
data = [e for e in data if e is not None]

need = [e for e in data if e.get('_referenznummer') and not e.get('E-Mail')]
print(f'Need details: {len(need)} employers', flush=True)

emails = 0; phones = 0; websites = 0

for i, emp in enumerate(need):
    refnr = emp.get('_referenznummer','')
    try:
        r = requests.get(DETAIL_API, headers=HEADERS, params={'refnr': refnr}, timeout=15)
        if r.status_code == 200:
            d = r.json()
            fd = d.get('firmaDetails', {})
            for p in fd.get('ansprechpartner', []):
                if p.get('email') and not emp.get('E-Mail'):
                    emp['E-Mail'] = p['email']; emails += 1
                if p.get('telefon') and not emp.get('Telefon'):
                    emp['Telefon'] = p['telefon']; phones += 1
            if fd.get('url') and not emp.get('Website'):
                emp['Website'] = fd['url']; websites += 1
            bw = d.get('bewerbung', {})
            if bw:
                if bw.get('email') and not emp.get('E-Mail'):
                    emp['E-Mail'] = bw['email']; emails += 1
                if bw.get('url') and not emp.get('Website'):
                    emp['Website'] = bw['url']; websites += 1
                if bw.get('telefon') and not emp.get('Telefon'):
                    emp['Telefon'] = bw['telefon']; phones += 1
    except:
        pass
    if (i+1) % 100 == 0:
        print(f'{i+1}/{len(need)} | E:{emails} P:{phones} W:{websites}', flush=True)
        with open(OUTPUT, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    time.sleep(random.uniform(0.1, 0.25))

with open(OUTPUT, 'w', encoding='utf-8') as f:
    json.dump(data, f, ensure_ascii=False, indent=2)

we = sum(1 for e in data if e.get('E-Mail'))
wp = sum(1 for e in data if e.get('Telefon'))
ww = sum(1 for e in data if e.get('Website'))
print(f'PHASE2 DONE: {len(data)} total | E:{we} P:{wp} W:{ww}', flush=True)
