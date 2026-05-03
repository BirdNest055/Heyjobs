#!/usr/bin/env python3
"""Phase 1: Collect all employer listings from Arbeitsagentur API"""
import os, re, json, time, random, hashlib, requests, sys
from datetime import datetime, timezone
sys.stdout.reconfigure(line_buffering=True)

API_BASE = 'https://rest.arbeitsagentur.de/jobboerse/jobsuche-service/pc/v6/jobs'
HEADERS = {
    'Accept': 'application/json, text/plain, */*',
    'X-API-Key': 'jobboerse-jobsuche',
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
    'Origin': 'https://www.arbeitsagentur.de', 'Referer': 'https://www.arbeitsagentur.de/',
}
OUTPUT = '/home/z/my-project/download/arbeitsagentur_employers.json'

with open(OUTPUT, 'r', encoding='utf-8') as f:
    employers = json.load(f)
employers = [e for e in employers if e is not None]
seen = set(e.get('_dedup_hash','') for e in employers if e.get('_dedup_hash'))
print(f'Loaded {len(employers)} employers', flush=True)

BL = {'BADEN_WUERTTEMBERG':'Baden-Württemberg','BAYERN':'Bayern','BERLIN':'Berlin','BRANDENBURG':'Brandenburg','BREMEN':'Bremen','HAMBURG':'Hamburg','HESSEN':'Hessen','MECKLENBURG_VORPOMMERN':'Mecklenburg-Vorpommern','NIEDERSACHSEN':'Niedersachsen','NORDRHEIN_WESTFALEN':'Nordrhein-Westfalen','RHEINLAND_PFALZ':'Rheinland-Pfalz','SAARLAND':'Saarland','SACHSEN':'Sachsen','SACHSEN_ANHALT':'Sachsen-Anhalt','SCHLESWIG_HOLSTEIN':'Schleswig-Holstein','THUERINGEN':'Thüringen'}

def gid(plz, name):
    c = re.sub(r'[^0-9]','',str(plz))[:5].zfill(5) if plz else '00000'
    h = int(hashlib.md5(f'{name}_{plz}'.lower().strip().encode()).hexdigest()[:4], 16) % 10000
    return f'DE-{c}-{h:04d}'

def proc(job):
    firma = job.get('firma','').strip()
    if not firma: return None
    locs = job.get('stellenlokationen',[])
    if not locs: return None
    a = locs[0].get('adresse',{})
    s,h,p,o,r = a.get('strasse',''),a.get('hausnummer',''),a.get('plz',''),a.get('ort',''),a.get('region','')
    dk = hashlib.md5(f'{firma.lower().strip()}_{p}'.encode()).hexdigest()
    if dk in seen: return None
    seen.add(dk)
    return {
        'Arbeitgeber_ID':gid(p,firma),'Name':firma,
        'Straße':f'{s} {h}'.strip() if s else '',
        'PLZ':p,'Stadt':o,'Bundesland':BL.get(r,r.replace('_',' ').title() if r else ''),
        'Telefon':'','E-Mail':'','Website':'',
        'Branche':'IT, Computer, Telekommunikation','Quelle':'Arbeitsagentur','Bewerbungsstatus':'Offen',
        '_referenznummer':job.get('referenznummer',''),
        '_ag_hash':job.get('arbeitgeberKundennummerHash',''),
        '_job_titel':job.get('stellenangebotsTitel',''),
        '_hauptberuf':job.get('hauptberuf',''),
        '_homeoffice':job.get('homeofficemoeglich',False),
        '_vertragsdauer':job.get('vertragsdauer',''),
        '_vollzeit':job.get('arbeitszeitVollzeit',False),
        '_teilzeit':bool(job.get('arbeitszeitTeilzeitVormittag') or job.get('arbeitszeitTeilzeitNachmittag')),
        '_ortsteil':a.get('ortsteil',''),'_dedup_hash':dk
    }

ub = datetime.now(timezone.utc).strftime('%Y-%m-%dT%H:%M:%S.000Z')
ce = 0
for pn in range(1, 351):
    try:
        r = requests.get(API_BASE, headers=HEADERS,
            params={'angebotsart':1,'branche':11,'page':pn,'size':25,'aktualisiertVor':ub,'pav':'false'}, timeout=30)
        if r.status_code == 200:
            el = r.json().get('ergebnisliste',[])
            if not el:
                ce += 1
                if ce >= 3: break
                continue
            ce = 0
            for j in el:
                e = proc(j)
                if e: employers.append(e)
        elif r.status_code == 429:
            time.sleep(15); continue
        else:
            ce += 1
            if ce >= 5: break
    except:
        ce += 1
        if ce >= 5: break
    if pn % 20 == 0:
        print(f'Page {pn}: {len(employers)} employers', flush=True)
        with open(OUTPUT,'w',encoding='utf-8') as f: json.dump(employers, f, ensure_ascii=False, indent=2)
    time.sleep(random.uniform(0.15, 0.4))

with open(OUTPUT,'w',encoding='utf-8') as f: json.dump(employers, f, ensure_ascii=False, indent=2)
print(f'PHASE1 DONE: {len(employers)} employers', flush=True)
