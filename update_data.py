"""
update_data.py v4
- Ek hi Excel file se dono sheets read karta hai:
  - Form sheet          → 10KW data
  - 5KW_INVOICE_DATA    → 5KW data
- data_v4.js + index.html rebuild karta hai
- GitHub Actions se har ghante run hota hai
"""
import os, json, requests, io
from datetime import datetime

MONTHS = ['January','February','March','April','May','June',
          'July','August','September','October','November','December']

ZONE_MAP = {
    'AGRA':'AGRA','AGRA ZONE':'AGRA','ALIGARH':'ALIGARH','ALIGARH ZONE':'ALIGARH',
    'BANDA':'BANDA','BANDA ZONE':'BANDA','ETAH':'ETAH','ETAH ZONE':'ETAH',
    'FIROZABAD':'FIROZABAD','FIROZABAD ZONE':'FIROZABAD','JHANSI':'JHANSI','JHANSI ZONE':'JHANSI',
    'KANPUR ZONE-1':'KANPUR-1','KANPUR-1':'KANPUR-1','KANPUR ZONE-2':'KANPUR-2','KANPUR-2':'KANPUR-2',
    'MATHURA':'MATHURA','MATHURA ZONE':'MATHURA',
}

print("=" * 55)
print("  DVVNL Dashboard — Auto Update v4")
print("=" * 55)

# ── 1. Download Excel file ────────────────────────────────────────────────────
SP_EXCEL = os.environ.get('SP_5KW', '')  # Same Excel for both sheets

if not SP_EXCEL:
    print("ERROR: SP_5KW secret not set!")
    exit(1)

url = SP_EXCEL + ('&' if '?' in SP_EXCEL else '?') + 'download=1'
print(f"\n[Excel] Downloading workbook...")
r = requests.get(url, timeout=60)
r.raise_for_status()
print(f"[Excel] Downloaded: {len(r.content)//1024} KB")

try:
    import openpyxl
    wb = openpyxl.load_workbook(io.BytesIO(r.content), read_only=True, data_only=True)
    print(f"[Excel] Sheets: {wb.sheetnames}")
except ImportError:
    print("ERROR: openpyxl not installed")
    exit(1)

# ── 2. Extract 10KW (Form sheet) ──────────────────────────────────────────────
print("\n[10KW] Reading Form sheet...")
ws_form = wb['Form']
records_10kw = []
for row in ws_form.iter_rows(min_row=2, values_only=True):
    if not row[0]: continue
    mon = str(row[6] or '').strip()
    yr  = str(row[7] or '').strip()
    if not mon or not yr: continue
    records_10kw.append({
        'zone':        str(row[10] or '').strip(),
        'verified_by': str(row[11] or '').strip(),
        'test_div':    str(row[12] or '').strip(),
        'dist_div':    str(row[13] or '').strip(),
        'circle':      str(row[14] or '').strip(),
        'month': mon, 'year': yr,
        'letter_no':   str(row[9]  or '').strip(),
    })
print(f"[10KW] Records: {len(records_10kw)}")

# ── 3. Extract 5KW (5KW_INVOICE_DATA sheet) ───────────────────────────────────
print("\n[5KW] Reading 5KW_INVOICE_DATA sheet...")
ws_5kw = wb['5KW_INVOICE_DATA']
records_5kw = []

def sf(v):
    try: return round(float(v), 2)
    except: return 0.0

for row in ws_5kw.iter_rows(min_row=3, values_only=True):
    zone_raw = str(row[0] or '').strip()
    zone = ZONE_MAP.get(zone_raw, '')
    if not zone: continue
    month_raw = row[5]
    if isinstance(month_raw, datetime):
        mon = MONTHS[month_raw.month-1]
        yr  = str(month_raw.year)
    else:
        continue
    records_5kw.append({
        'zone': zone, 'month': mon, 'year': yr,
        'miro':      str(row[2]  or '').strip(),
        'invoice':   str(row[3]  or '').strip(),
        'basic':     sf(row[11]),
        'gst':       sf(row[12]),
        'total':     sf(row[13]),
        'status_70': str(row[14] or '').strip(),
        'status_30': str(row[16] or '').strip(),
        'from_zone': str(row[20] or '').strip(),
        'submit':    str(row[21] or '').strip(),
    })
print(f"[5KW] Records: {len(records_5kw)}")

# ── 4. Master data ────────────────────────────────────────────────────────────
print("\n[Master] Loading master_data.json...")
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
print(f"[Master] Divisions: {len(MASTER_ALL)}")

# ── 5. Compute 10KW monthly status ───────────────────────────────────────────
print("\n[10KW] Computing monthly status...")
all_months = sorted(
    set((r['month'], r['year']) for r in records_10kw),
    key=lambda x: (int(x[1]), MONTHS.index(x[0])) if x[0] in MONTHS else (0,0)
)
MONTHLY_STATUS = {}
for mon, yr in all_months:
    recs = [r for r in records_10kw if r['month']==mon and r['year']==yr]
    vd = set(r['dist_div'].upper().strip() for r in recs if r['verified_by']=='EX EN DISTRIBUTION' and r['dist_div'])
    vt = set(r['test_div'].upper().strip() for r in recs if r['verified_by']=='EX EN TEST' and r['test_div'])
    vc = set(r['circle'].upper().strip() for r in recs if r['verified_by']=='CIRCLE OFFICE - SE' and r['circle'])
    result = []
    for m in MASTER_ALL:
        d = m['division'].upper().strip()
        if   m['type']=='DIST':   status = 'YES' if d in vd else 'NO'
        elif m['type']=='TEST':   status = 'YES' if d in vt else 'NO'
        else:                     status = 'YES' if d in vc else 'NO'
        result.append({'zone':m['zone'],'circle':m['circle'],
                       'type':m['type'],'division':m['division'],'status':status})
    MONTHLY_STATUS[f"{mon}_{yr}"] = result
    yes = sum(1 for r in result if r['status']=='YES')
    print(f"  {mon:12} {yr}: {yes}/{len(result)} ({round(yes/len(result)*100)}%)")

mini_10kw = [[r['zone'],r['verified_by'],r['test_div'],r['dist_div'],
              r['circle'],r['month'],r['year'],r['letter_no']] for r in records_10kw]

# ── 6. Write data_v4.js ───────────────────────────────────────────────────────
print("\n[Build] Writing data_v4.js...")
ts = datetime.now().strftime('%d %b %Y %H:%M')
data_js = (
    f"// Auto-updated: {ts} | 10KW: {len(records_10kw)} | 5KW: {len(records_5kw)} | Months: {len(MONTHLY_STATUS)}\n"
    f"const MASTER_ALL={json.dumps(MASTER_ALL, separators=(',',':'))};\n"
    f"const MONTHLY_STATUS={json.dumps(MONTHLY_STATUS, separators=(',',':'))};\n"
    f"const FORM_RECORDS={json.dumps(mini_10kw, separators=(',',':'))};\n"
    f"var RECORDS_5KW={json.dumps(records_5kw, separators=(',',':'))};\n"
)
with open('data_v4.js', 'w', encoding='utf-8') as f:
    f.write(data_js)
print(f"[Build] data_v4.js: {os.path.getsize('data_v4.js')//1024} KB")

# ── 7. Rebuild index.html ─────────────────────────────────────────────────────
print("[Build] Rebuilding index.html...")
with open('index_template.html', encoding='utf-8') as f:
    template = f.read()
index_html = template.replace('__DATA_PLACEHOLDER__', data_js)
with open('index.html', 'w', encoding='utf-8') as f:
    f.write(index_html)
print(f"[Build] index.html: {os.path.getsize('index.html')//1024} KB")

print(f"\n{'='*55}")
print(f"  ✅ Done! Updated at {ts}")
print(f"  10KW: {len(MONTHLY_STATUS)} months | 5KW: {len(records_5kw)} records")
print(f"{'='*55}\n")
