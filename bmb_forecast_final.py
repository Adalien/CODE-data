# -*- coding: utf-8 -*-
import pandas as pd, warnings, os, glob, sys, io, re, math
warnings.filterwarnings('ignore')
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

BASE = r"C:\Users\admin\OneDrive\桌面\CODE資料"
DESK = r"C:\Users\admin\OneDrive\桌面"

def latest2(*pats):
    files = []
    for pat in pats:
        files.extend(glob.glob(pat))
    return max(files, key=os.path.getmtime) if files else None

order_path  = latest2(os.path.join(DESK, "訂單生產總表*.xlsx"), os.path.join(BASE, "訂單生產總表*.xlsx"))
master_path = latest2(os.path.join(BASE, "BOM*.xlsx"))
inv_path    = latest2(os.path.join(DESK, "原物料庫存*.xlsx"), os.path.join(BASE, "原物料庫存*.xlsx"))

print(f"訂單: {os.path.basename(order_path)}")
print(f"主BOM: {os.path.basename(master_path)}")
print(f"庫存: {os.path.basename(inv_path)}")
print()

def get_roll_len(mat_code):
    """從品號末段取每捲米數，BMB 預設 1500"""
    code = mat_code.replace('-','')
    m = re.search(r'(\d{3,4})$', code)
    if m:
        cand = int(m.group(1))
        if 500 <= cand <= 9999:
            return cand
    return 1500  # BMB 預設

# ── 讀主BOM 熔噴用量 ──────────────────────────────────────
master = pd.read_excel(master_path)
master.columns = [c.strip() for c in master.columns]
bmb = master[master['元件品號'].astype(str).str.startswith('BMB')].copy()
bmb = bmb[bmb['主件品號'].notna() & (bmb['主件品號'].astype(str) != 'nan')].copy()
bmb['組成用量'] = pd.to_numeric(bmb['組成用量'], errors='coerce').fillna(0)
bmb['底數']    = pd.to_numeric(bmb['底數'],    errors='coerce').replace(0, 1).fillna(1)
bmb['損耗率%'] = pd.to_numeric(bmb['損耗率%'], errors='coerce').fillna(0)
# 每片需求公尺數
bmb['m_per_片'] = bmb['組成用量'] / bmb['底數'] * (1 + bmb['損耗率%'] / 100)
# 每捲長度
bmb['捲長m'] = bmb['元件品號'].apply(get_roll_len)
# 每片需求捲數
bmb['捲_per_片'] = bmb['m_per_片'] / bmb['捲長m']

bmb_map = bmb[['元件品號','主件品號','捲_per_片','m_per_片','捲長m','品     名']].rename(
    columns={'元件品號':'BMB品號','主件品號':'成品品號','品     名':'BMB品名'})

# ── 讀訂單 ──────────────────────────────────────────────
df = pd.read_excel(order_path, sheet_name='總表', header=0)
df.columns = ['訂單日期','序列','通路','品項','品號','品名','單位','機台','生產時間',
              '訂單量','庫存','差異','生產量','箱數','批號','製令單號','領料','入庫',
              '交期','出貨日','備註','機台1','生產時間1'][:len(df.columns)]

df = df[df['入庫'].isna() | (df['入庫'] == 0)]
df = df[df['領料'].isna() | (df['領料'] == 0)]
df['生產量'] = pd.to_numeric(df['生產量'], errors='coerce').fillna(0)
df['訂單量'] = pd.to_numeric(df['訂單量'], errors='coerce').fillna(0)
df['出貨日'] = pd.to_datetime(df['出貨日'], errors='coerce')
df['月份']   = df['出貨日'].dt.month
df['品號']   = df['品號'].astype(str).str.strip()
df['需產量'] = df['生產量'].where(df['生產量'] > 0, df['訂單量'])

target = df[df['月份'].isin([5, 6])].copy()
print(f"5-6月未完成訂單: {len(target)} 筆")
print(f"  5月: {len(target[target['月份']==5])} 筆  需產量: {target[target['月份']==5]['需產量'].sum():,.0f} 片")
print(f"  6月: {len(target[target['月份']==6])} 筆  需產量: {target[target['月份']==6]['需產量'].sum():,.0f} 片")

# ── Merge 訂單 x BOM ──────────────────────────────────────
merged = target.merge(bmb_map, left_on='品號', right_on='成品品號', how='left')
matched = merged[merged['BMB品號'].notna()].copy()
matched['需求捲數'] = (matched['需產量'] * matched['捲_per_片']).round(1)

# ── 讀庫存（從缺料分析抓BMB捲數）──────────────────────────
inv_raw = pd.read_excel(inv_path, header=None)
inv_dict = {}
for _, row in inv_raw.iterrows():
    for i, cell in enumerate(row):
        c = str(cell).strip()
        if c.startswith('BMB'):
            for j in range(i+1, min(i+6, len(row))):
                try:
                    qty = float(row[j])
                    if qty > 0:
                        inv_dict[c] = inv_dict.get(c, 0) + qty
                        break
                except:
                    continue

# ── 輸出結果 ──────────────────────────────────────────────
print()
print("="*65)
print("5-6月 熔噴布需求預估（依品號）")
print("="*65)

grand_total = 0
for month, label in [(5,'5月'), (6,'6月')]:
    sub = matched[matched['月份'] == month]
    s = sub.groupby(['BMB品號','捲長m']).agg(
        需求捲數=('需求捲數','sum'),
    ).reset_index().sort_values('需求捲數', ascending=False)
    total = s['需求捲數'].sum()
    grand_total += total
    print(f"\n【{label}】  小計: {total:.1f} 捲")
    print(f"  {'BMB品號':<40} {'捲長':>5} {'需求捲':>8}")
    print(f"  {'-'*40} {'-'*5} {'-'*8}")
    for _, r in s.iterrows():
        print(f"  {r['BMB品號']:<40} {int(r['捲長m']):>5} {r['需求捲數']:>8.1f}")

print()
print("="*65)
print("庫存 vs 5-6月總需求（缺口分析）")
print("="*65)

all_need = matched.groupby(['BMB品號','捲長m'])['需求捲數'].sum().reset_index()
all_need = all_need.sort_values('需求捲數', ascending=False)

# 現有庫存（缺料分析取到的捲數）
# 這裡直接用熔噴現貨品號
BMB_INV = {
    'BMB-MG-H11W201751500': 6.0,
    'BMB-MG-H11W202601500': 52.0,
    'BMB-MG-H12D252601200': 0.0,
    'BMB-MG-H12S252601200': 0.0,
    'BMB-MG-H12W252601500': 0.0,
    'BMB-MG-MEDG201751500': 30.0,
    'BMB-MG-MEDW201751500': 10.0,
    'BMB-MG-MEDG202601500': 8.0,
    'BMB-MG-MEDW202601500': 42.0,
    'BMB-MG250-DGY020-H11': 57.0,
    'BMB-MG250-LGY020-H11': 57.0,
    'BMB-MG-H11G202501500': 57.0,
    'BMB-MG-H11S202501500': 57.0,
    'BMB-MT-E12W252501000': 0.0,
}

print(f"\n  {'BMB品號':<40} {'現貨':>6} {'需求':>7} {'缺口':>7}")
print(f"  {'-'*40} {'-'*6} {'-'*7} {'-'*7}")
total_need = 0
total_gap  = 0
for _, r in all_need.iterrows():
    bmb_id = r['BMB品號']
    need   = r['需求捲數']
    stock  = BMB_INV.get(bmb_id, 0)
    gap    = max(0, need - stock)
    total_need += need
    total_gap  += gap
    flag = "❌" if gap > 0 else "✓ "
    print(f"  {flag} {bmb_id:<40} {stock:>6.1f} {need:>7.1f} {gap:>7.1f}")

print(f"\n  {'合計':<42} {sum(BMB_INV.values()):>6.1f} {total_need:>7.1f} {total_gap:>7.1f}")
print(f"\n  5-6月熔噴總需求: {total_need:.1f} 捲")
print(f"  現有庫存合計:    {sum(BMB_INV.values()):.1f} 捲")
print(f"  需補購:          {total_gap:.1f} 捲")
