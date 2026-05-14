#!/usr/bin/env python3
"""Download HTML for websites that were found by search but not yet downloaded."""
import json, re, time, os, sys, socket
from datetime import datetime
import urllib.request, urllib.parse, urllib.error

ENRICH_FILE = '/home/z/my-project/website_enrichment.json'
HTML_DIR = '/home/z/my-project/download/website_html'

def safe_filename(name):
    name = re.sub(r'[^\w\s\-.]', '', name)
    name = re.sub(r'\s+', '_', name.strip())
    return name[:80]

def url_get(url, timeout=15):
    try:
        req = urllib.request.Request(url, headers={
            'User-Agent': 'Mozilla/5.0 (X11; Linux x86_64; rv:124.0) Gecko/20100101 Firefox/124.0',
            'Accept': 'text/html,application/xhtml+xml,*/*',
            'Accept-Language': 'de-DE,de;q=0.9',
        })
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            return resp.read(), resp.url, resp.headers.get('Content-Type','')
    except:
        return None, None, None

def extract_emails(html_bytes):
    try: text = html_bytes.decode('utf-8', errors='replace')
    except: text = html_bytes.decode('latin-1', errors='replace')
    emails = set()
    for m in re.finditer(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}', text):
        e = m.group(0).lower()
        if not any(x in e for x in ['.png','.jpg','.css','.js','example.com','sentry','wixpress']):
            emails.add(e)
    for m in re.finditer(r'mailto:([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})', text, re.I):
        emails.add(m.group(1).lower())
    skip = ['wordpress','admin@localhost','noreply@','webmaster@']
    emails = {e for e in emails if not any(s in e for s in skip)}
    pref = [e for e in emails if any(k in e for k in ['bewerb','hr','karrier','job'])]
    if pref: return pref[0]
    pref2 = [e for e in emails if any(k in e for k in ['info','kontakt','contact'])]
    if pref2: return pref2[0]
    return sorted(emails)[0] if emails else ''

def try_impressum(base_url, safe_name):
    for path in ['/impressum','/Impressum','/kontakt','/Kontakt','/contact']:
        url = base_url.rstrip('/') + path
        data, _, ct = url_get(url, timeout=8)
        if data and 'text/html' in (ct or ''):
            email = extract_emails(data)
            if email:
                fp = os.path.join(HTML_DIR, f"{safe_name}_imp.html")
                with open(fp, 'wb') as f: f.write(data)
                return email, url
    return '', ''

def main():
    with open(ENRICH_FILE) as f: enrichment = json.load(f)
    
    need = {k:v for k,v in enrichment.items() if v.get('website') and not v.get('html_file') and v.get('status') != 'no_website'}
    print(f"Need HTML download: {len(need)}")
    
    ok = fail = emails = 0
    for i, (name, data) in enumerate(need.items()):
        website = data['website']
        safe = safe_filename(name)
        html_path = os.path.join(HTML_DIR, f"{safe}.html")
        rel_path = f"website_html/{safe}.html"
        
        print(f"[{i+1}/{len(need)}] {name}: {website[:60]}", flush=True)
        
        content, final_url, ct = url_get(website)
        
        email = ''
        imp_url = ''
        
        if content and 'text/html' in (ct or ''):
            with open(html_path, 'wb') as f: f.write(content)
            email = extract_emails(content)
            if not email:
                email, imp_url = try_impressum(final_url or website, safe)
            ok += 1
            if email: emails += 1
            print(f"  OK {len(content)}b email={'YES' if email else '-'}", flush=True)
            enrichment[name].update({
                'html_file': rel_path, 'email': email,
                'impressum_url': imp_url, 'html_size': len(content),
                'status': 'success'
            })
        else:
            fail += 1
            print(f"  FAIL", flush=True)
            enrichment[name]['status'] = 'html_fail'
        
        with open(ENRICH_FILE,'w') as f: json.dump(enrichment,f,ensure_ascii=False)
        time.sleep(1)
    
    print(f"\nDone! OK:{ok} Fail:{fail} Emails:{emails}")

if __name__ == '__main__':
    main()
