"""
update_data.py v2
- SharePoint se formdata.json fetch karta hai
- data_v4.js regenerate karta hai
- index.html bhi rebuild karta hai (data embedded)
GitHub Actions se har ghante run hota hai.
"""
import os, json, requests
from datetime import datetime

MONTHS = ['January','February','March','April','May','June',
          'July','August','September','October','November','December']

# ── 1. Fetch data ─────────────────────────────────────────────────────────────
SP_URL = os.environ.get('SP_URL', '')

if SP_URL:
    url = SP_URL + ('&' if '?' in SP_URL else '?') + 'download=1'
    print(f"Fetching SharePoint data...")
    r = requests.get(url, timeout=30)
    r.raise_for_status()
    raw = r.json()
    print(f"Fetched {len(raw)} records")
else:
    print("SP_URL not set — using local formdata.json")
    with open('formdata.json') as f:
        raw = json.load(f)
    print(f"Local records: {len(raw)}")

# ── 2. Clean records ──────────────────────────────────────────────────────────
records = []
for r in raw:
    mon = str(r.get('Month Of Letter') or r.get('month') or '').strip()
    yr  = str(r.get('Year of Letter')  or r.get('year')  or '').strip()
    if not mon or not yr: continue
    records.append({
        'zone':        str(r.get('ZONE NAME')                     or r.get('zone','')       ).strip(),
        'verified_by': str(r.get('Invoice Letter Verified BY')    or r.get('verified_by','') ).strip(),
        'test_div':    str(r.get('NAME OF TEST DIVISION')         or r.get('test_div','')    ).strip(),
        'dist_div':    str(r.get('NAME OF DISTRIBUTOIN DIVISION') or r.get('dist_div','')    ).strip(),
        'circle':      str(r.get('NAME CIRLCE OFFICE (SE)')       or r.get('circle','')      ).strip(),
        'month': mon, 'year': yr,
        'letter_no':   str(r.get('Letter Number ')                or r.get('letter_no','')   ).strip(),
    })
print(f"Clean records: {len(records)}")

# ── 3. Master data ────────────────────────────────────────────────────────────
with open('master_data.json') as f:
    master = json.load(f)

master_dist = [{'zone':m['zone'],'circle':m['circle'],'test':m['test'],
                'division':m['dist'],'type':'DIST'} for m in master]
seen_t, seen_c = set(), set()
master_test, master_circle = [], []
for m in master:
    if m['test'] not in seen_t:
        seen_t.add(m['test'])
        master_test.append({'zone':m['zone'],'circle':m['circle'],
                            'division':m['test'],'type':'TEST'})
    if m['circle'] not in seen_c:
        seen_c.add(m['circle'])
        master_circle.append({'zone':m['zone'],'circle':m['circle'],
                              'division':m['circle'],'type':'CIRCLE'})
MASTER_ALL = master_dist + master_test + master_circle

# ── 4. Compute monthly status ─────────────────────────────────────────────────
all_months = sorted(
    set((r['month'], r['year']) for r in records),
    key=lambda x: (int(x[1]), MONTHS.index(x[0])) if x[0] in MONTHS else (0,0)
)
MONTHLY_STATUS = {}
for mon, yr in all_months:
    recs = [r for r in records if r['month']==mon and r['year']==yr]
    vd = set(r['dist_div'].upper().strip() for r in recs if r['verified_by']=='EX EN DISTRIBUTION' and r['dist_div'])
    vt = set(r['test_div'].upper().strip() for r in recs if r['verified_by']=='EX EN TEST' and r['test_div'])
    vc = set(r['circle'].upper().strip() for r in recs if r['verified_by']=='CIRCLE OFFICE - SE' and r['circle'])
    result = []
    for m in MASTER_ALL:
        d = m['division'].upper().strip()
        if   m['type']=='DIST':   status='YES' if d in vd else 'NO'
        elif m['type']=='TEST':   status='YES' if d in vt else 'NO'
        else:                     status='YES' if d in vc else 'NO'
        result.append({'zone':m['zone'],'circle':m['circle'],
                       'type':m['type'],'division':m['division'],'status':status})
    MONTHLY_STATUS[f"{mon}_{yr}"] = result
    yes = sum(1 for r in result if r['status']=='YES')
    print(f"  {mon} {yr}: {yes}/{len(result)} ({round(yes/len(result)*100)}%)")

mini = [[r['zone'],r['verified_by'],r['test_div'],r['dist_div'],
         r['circle'],r['month'],r['year'],r['letter_no']] for r in records]

# ── 5. Write data_v4.js ───────────────────────────────────────────────────────
ts = datetime.now().strftime('%d %b %Y %H:%M')
data_js = (
    f"// Auto-updated: {ts} | Records: {len(records)} | Months: {len(MONTHLY_STATUS)}\n"
    f"const MASTER_ALL={json.dumps(MASTER_ALL, separators=(',',':'))};\n"
    f"const MONTHLY_STATUS={json.dumps(MONTHLY_STATUS, separators=(',',':'))};\n"
    f"const FORM_RECORDS={json.dumps(mini, separators=(',',':'))};\n"
)
with open('data_v4.js','w',encoding='utf-8') as f:
    f.write(data_js)
print(f"data_v4.js: {os.path.getsize('data_v4.js')//1024} KB")

# ── 6. Rebuild index.html with fresh data ────────────────────────────────────
with open('index_template.html', encoding='utf-8') as f:
    template = f.read()

index_html = template.replace('__DATA_PLACEHOLDER__', data_js)
with open('index.html','w',encoding='utf-8') as f:
    f.write(index_html)
print(f"index.html rebuilt: {os.path.getsize('index.html')//1024} KB")
print(f"Done! Updated at {ts}")
