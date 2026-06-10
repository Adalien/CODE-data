# -*- coding: utf-8 -*-
"""
gen_html_v2.py — 模板 + 資料分離版
讀 進貨明細_template.html → 填入資料 → 輸出 HTML
★ 完全不用 regex 操作 JS/CSS，只做 str.replace()
"""
import pandas as pd
import json, glob, os, sys, base64, re
from datetime import datetime, date
sys.stdout.reconfigure(encoding='utf-8')

# ── 路徑 ──────────────────────────────────────────────────────
BASE    = r'C:\Users\admin\OneDrive\桌面\CODE資料'
DESKTOP = r'C:\Users\admin\OneDrive\桌面'
TMPL    = os.path.join(BASE, '進貨明細_template.html')
OUTPUTS = [
    os.path.join(DESKTOP, '進貨明細查詢.html'),
    os.path.join(BASE,    '進貨明細查詢.html'),
    os.path.join(DESKTOP, '達特世生技_進貨明細查詢_2026.html'),
    os.path.join(BASE,    '達特世生技_進貨明細查詢_2026.html'),
]

# ── 工具函式 ──────────────────────────────────────────────────
def esc(v):
    if v is None: return ''
    return str(v).replace('&','&amp;').replace('<','&lt;').replace('>','&gt;')

def fnum(v):
    if v is None: return ''
    try:
        if pd.isna(v): return ''
        n = float(v)
        if n == int(n): return f'{int(n):,}'
        return f'{n:,.2f}'.rstrip('0').rstrip('.')
    except: return esc(str(v))

def clean_id(v):
    if v is None: return ''
    try:
        if pd.isna(v): return ''
    except: pass
    s = str(v).strip()
    if s in ('nan', 'None', ''): return ''
    if s.endswith('.0') and s[:-2].replace('-','').isdigit(): return s[:-2]
    return s

def ds(v):
    try:
        if pd.isna(v): return ''
    except: pass
    try: return pd.to_datetime(v).strftime('%Y/%m/%d')
    except:
        s = str(v).strip()
        return '' if s in ('nan','NaT','None','') else s

def src_div(fname, date_str):
    return f'<div class="data-source">資料來源：{esc(fname)}（{date_str}）</div>\n'

# ── 輪播圖片 ──────────────────────────────────────────────────
def load_slideshow():
    """從 輪播圖片/ 資料夾載入 base64 圖片，回傳 (imgs_html, nimgs)"""
    MIME = {'.jpg':'image/jpeg','.jpeg':'image/jpeg',
            '.gif':'image/gif','.png':'image/png','.webp':'image/webp'}
    folder = os.path.join(BASE, '輪播圖片')
    imgs = []
    if os.path.isdir(folder):
        for fn in sorted(os.listdir(folder)):
            ext = os.path.splitext(fn)[1].lower()
            if ext not in MIME: continue
            with open(os.path.join(folder, fn), 'rb') as f:
                imgs.append(f'data:{MIME[ext]};base64,{base64.b64encode(f.read()).decode()}')
        print(f'  輪播：{len(imgs)} 張（{", ".join(sorted(os.listdir(folder)))}）')
    if not imgs:
        print('  [WARN] 找不到輪播圖片，使用空白')
        return '', 0
    IS = 'position:absolute;inset:0;width:100%;height:100%;object-fit:contain;transition:opacity 0.3s;'
    html = ''.join(
        f'<img src="{s}" style="{IS}opacity:{"1" if i==0 else "0"};">'
        for i, s in enumerate(imgs)
    )
    return html, len(imgs)

def mascot_div(mid, imgs_html):
    """產生一個 mascot filter-card（放在每個分頁頂端）"""
    return (
        f'<div class="filter-card" style="display:flex;align-items:flex-end;">'
        f'<div style="flex:1;"></div>'
        f'<div id="{mid}" style="flex:0 0 auto;width:260px;height:110px;'
        f'position:relative;border-radius:8px;overflow:hidden;">'
        f'{imgs_html}</div>'
        f'</div>\n'
    )

# ── 各頁籤 HTML 產生函式 ───────────────────────────────────────
def gen_orders(path, fname, date_str, imgs_html):
    try:
        df = pd.read_excel(path, sheet_name='訂單交期總覽', header=1)
    except Exception as e:
        return mascot_div('mascot-tab-orders', imgs_html) + f'<div style="padding:20px;color:#888">找不到訂單交期總覽：{e}</div>'
    df.columns = [str(c).strip() for c in df.columns]
    today = pd.Timestamp(date.today())
    rows_html = ''; data_count = 0
    for _, r in df.iterrows():
        pno = str(r.get('品號','')).strip()
        if not pno or pno in ('nan','None',''): continue
        data_count += 1
        try:
            dl_ts = pd.to_datetime(r.get('交期','')) if str(r.get('交期','')).strip() not in ('','nan','NaT') else None
        except: dl_ts = None
        cls = 'lack-severe' if dl_ts and (dl_ts - today).days <= 7 else ''
        _pno = esc(pno)
        rows_html += (
            f'<tr class="{cls}">'
            f'<td>{ds(r.get("訂單日期",""))}</td>'
            f'<td>{esc(r.get("序列",""))}</td>'
            f'<td>{esc(r.get("通路",""))}</td>'
            f'<td>{esc(r.get("品項",""))}</td>'
            f'<td class="ord-pno" style="font-size:12px" onclick="showProductBOM(\'{_pno}\')">{_pno}</td>'
            f'<td style="font-size:12px">{esc(r.get("品名",""))}</td>'
            f'<td>{esc(r.get("機台",""))}</td>'
            f'<td>{esc(r.get("生產時間",""))}</td>'
            f'<td class="num">{fnum(r.get("訂單量",""))}</td>'
            f'<td class="num">{fnum(r.get("生產量",""))}</td>'
            f'<td style="font-size:12px">{clean_id(r.get("批號",""))}</td>'
            f'<td style="font-size:12px">{clean_id(r.get("製令單號",""))}</td>'
            f'<td>{ds(r.get("領料",""))}</td>'
            f'<td>{ds(r.get("交期",""))}</td>'
            f'<td>{ds(r.get("出貨日",""))}</td>'
            f'<td style="font-size:12px">{esc(r.get("備註",""))}</td>'
            f'</tr>\n'
        )
    filter_row = (
        '<tr class="col-filter-row">'
        + ''.join(
            f'<th><input class="col-filter ord-col-filter" data-colidx="{i}"'
            f' placeholder="篩選" type="text" oninput="ordFilterInput(this)"></th>'
            for i in range(16)
        ) + '</tr>'
    )
    stats = (
        f'<div class="shortage-stats">'
        f'<span style="font-size:13px;color:#4a5568">共 <strong style="color:#1a56a0">{data_count}</strong> 筆待生產訂單</span>'
        f'<span class="stat-badge stat-red" style="margin-left:12px">🔴 交期7天內</span>'
        f'</div>\n'
    )
    table = (
        f'<div class="table-wrap" style="margin:10px 18px 24px">'
        f'<div class="table-scroll"><table><thead><tr>'
        f'<th>訂單日期</th><th>序列</th><th>通路</th><th>品項</th>'
        f'<th>品號</th><th>品名</th><th>機台</th><th>生產時間</th>'
        f'<th class="num">訂單量</th><th class="num">生產量</th>'
        f'<th>批號</th><th>製令單號</th><th>領料</th><th>交期</th><th>出貨日</th><th>備註</th>'
        f'</tr>{filter_row}</thead><tbody>\n{rows_html}</tbody></table></div></div>'
    )
    return mascot_div('mascot-tab-orders', imgs_html) + src_div(fname, date_str) + stats + table


def gen_shortage(path, fname, date_str, imgs_html):
    df = pd.read_excel(path, sheet_name='缺料清單', header=1)
    if len(df.columns) >= 13:
        df.columns = ['材料類型','材料品號','品名(系統)','材料品名','庫存對應品項',
                      '需備量','單位','現有庫存','缺少量','急迫程度','訂單量(盒)','訂單筆數','最早交期']
    else:
        df.columns = ['材料類型','材料品號','品名(系統)','材料品名','庫存對應品項',
                      '需備量','單位','現有庫存','缺少量','急迫程度','訂單量(盒)','訂單筆數']
    counts = {'🔴':0,'🟠':0,'🟡':0,'📋':0}
    rows_html = ''
    for _, r in df.iterrows():
        mat = str(r['材料類型']) if pd.notna(r['材料類型']) else ''
        if not mat or mat == 'nan': continue
        if str(mat).startswith('▌') or str(mat).startswith('【'):
            rows_html += f'<tr class="section-header"><td colspan="9">{esc(mat)}</td></tr>\n'
            continue
        urg = str(r['急迫程度']) if pd.notna(r['急迫程度']) else ''
        if '🔴' in urg:   counts['🔴'] += 1; cls = 'lack-severe'
        elif '🟠' in urg: counts['🟠'] += 1; cls = 'lack-high'
        elif '🟡' in urg: counts['🟡'] += 1; cls = 'lack-medium'
        else:              counts['📋'] += 1; cls = ''
        rows_html += (
            f'<tr class="{cls}">'
            f'<td>{esc(mat)}</td>'
            f'<td>{esc(r["材料品號"])}</td>'
            f'<td>{esc(r["材料品名"] if pd.notna(r["材料品名"]) else r["品名(系統)"])}</td>'
            f'<td class="num">{fnum(r["需備量"])}</td>'
            f'<td class="num">{fnum(r["現有庫存"])}</td>'
            f'<td class="num">{fnum(r["缺少量"])}</td>'
            f'<td>{esc(urg)}</td>'
            f'<td class="num">{fnum(r["訂單量(盒)"])}</td>'
            f'<td class="num">{fnum(r["訂單筆數"])}</td>'
            f'</tr>\n'
        )
    total = sum(counts.values())
    stats = (
        f'<div class="shortage-stats">\n'
        f'<span class="stat-badge stat-red">🔴 嚴重 {counts["🔴"]} 項</span>\n'
        f'<span class="stat-badge stat-orange">🟠 高 {counts["🟠"]} 項</span>\n'
        f'<span class="stat-badge stat-yellow">🟡 中 {counts["🟡"]} 項</span>\n'
        f'<span class="stat-badge stat-blue">📋 需叫布 {counts["📋"]} 項</span>\n'
        f'<span style="font-size:13px;color:#718096">合計 {total} 項缺料</span>\n'
        f'</div>\n'
    )
    table = (
        f'<div class="table-wrap" style="margin:10px 18px 24px">'
        f'<div class="table-scroll"><table><thead><tr>'
        f'<th>材料類型</th><th>材料品號</th><th>材料品名</th>'
        f'<th class="num">需備量</th><th class="num">現有庫存</th><th class="num">缺少量</th>'
        f'<th>急迫程度</th><th class="num">訂單量(盒)</th><th class="num">訂單筆數</th>'
        f'</tr></thead><tbody>\n{rows_html}</tbody></table></div></div>'
    )
    return mascot_div('mascot-tab-shortage', imgs_html) + src_div(fname, date_str) + stats + table


def gen_inventory(path, fname, date_str, imgs_html):
    df = pd.read_excel(path, sheet_name='庫存解析驗證', header=1)
    df.columns = ['類別','庫存原始字串','解析量','單位','備註']
    rows_html = ''
    for _, r in df.iterrows():
        cat = str(r['類別']) if pd.notna(r['類別']) else ''
        if not cat or cat == 'nan': continue
        if str(cat).startswith('【'):
            rows_html += f'<tr class="cat-header"><td colspan="5">{esc(cat)}</td></tr>\n'; continue
        orig = str(r['庫存原始字串']) if pd.notna(r['庫存原始字串']) else ''
        qty = r['解析量']; unit = str(r['單位']) if pd.notna(r['單位']) else ''
        note = str(r['備註']) if pd.notna(r['備註']) else ''
        is_zero = False
        try: is_zero = float(qty) == 0
        except: pass
        rows_html += (
            f'<tr class="{"inv-zero" if is_zero else ""}">'
            f'<td>{esc(cat)}</td><td style="font-size:13px">{esc(orig)}</td>'
            f'<td class="num">{fnum(qty)}</td><td>{esc(unit)}</td>'
            f'<td style="font-size:12px">{esc(note)}</td></tr>\n'
        )
    table = (
        f'<div class="table-wrap" style="margin:10px 18px 24px">'
        f'<div class="table-scroll"><table>'
        f'<thead><tr><th>類別</th><th>庫存原始字串</th>'
        f'<th class="num">解析量</th><th>單位</th><th>備註</th></tr></thead>'
        f'<tbody>\n{rows_html}</tbody></table></div></div>'
    )
    return mascot_div('mascot-tab-inventory', imgs_html) + src_div(fname, date_str) + table


def gen_material(path, fname, date_str, imgs_html):
    df = pd.read_excel(path, sheet_name='材料需求彙總', header=0)
    df.columns = [str(c).strip() for c in df.columns]
    lack_n = (df['狀態'].str.contains('缺料', na=False)).sum()
    ok_n   = len(df) - lack_n
    unit_col  = '單位'
    unit1_col = '單位.1' if '單位.1' in df.columns else '單位'
    rows_html = ''
    for _, r in df.iterrows():
        status = str(r['狀態']) if pd.notna(r.get('狀態')) else ''
        rows_html += (
            f'<tr class="{"lack-severe" if "缺料" in status else "inv-ok"}">'
            f'<td>{esc(r["材料類型"])}</td>'
            f'<td>{esc(r["材料品號"])}</td>'
            f'<td style="font-size:12px">{esc(r.get("品名(系統)",""))}</td>'
            f'<td>{esc(r.get("材料品名","") if pd.notna(r.get("材料品名")) else "")}</td>'
            f'<td class="num">{fnum(r.get("總需求量",""))}</td>'
            f'<td>{esc(r.get(unit_col,""))}</td>'
            f'<td class="num">{fnum(r.get("庫存量",""))}</td>'
            f'<td>{esc(r.get(unit1_col,""))}</td>'
            f'<td class="num">{fnum(r.get("淨缺量",""))}</td>'
            f'<td>{esc(status)}</td></tr>\n'
        )
    stats = (
        f'<div class="shortage-stats">\n'
        f'<span style="font-size:13px;color:#4a5568">總品項 <strong style="color:#1a56a0">{len(df)}</strong></span>\n'
        f'<span class="stat-badge stat-red">⚠ 缺料 {lack_n}</span>\n'
        f'<span class="stat-badge" style="background:#f0fff4;color:#276749">✓ 充足 {ok_n}</span>\n'
        f'</div>\n'
    )
    table = (
        f'<div class="table-wrap" style="margin:10px 18px 24px">'
        f'<div class="table-scroll"><table><thead><tr>'
        f'<th>材料類型</th><th>材料品號</th><th style="font-size:12px">品名(系統)</th>'
        f'<th>材料品名</th><th class="num">總需求量</th><th>單位</th>'
        f'<th class="num">庫存量</th><th>單位</th><th class="num">淨缺量</th><th>狀態</th>'
        f'</tr></thead><tbody>\n{rows_html}</tbody></table></div></div>'
    )
    return mascot_div('mascot-tab-material', imgs_html) + src_div(fname, date_str) + stats + table


def gen_pick(path, fname, date_str, imgs_html):
    df = pd.read_excel(path, sheet_name='須領料', header=0)
    ncols = len(df.columns)
    if ncols >= 11:
        df.columns = ['品項','品號','品名','機台','生產時間','出貨日','交期','訂單量','差異(生產量)','製令單號','批號']
        has_ship = True; has_dl = True
    elif ncols >= 10:
        df.columns = ['品項','品號','品名','機台','生產時間','交期','訂單量','差異(生產量)','製令單號','批號']
        has_ship = False; has_dl = True
    else:
        df.columns = ['品項','品號','品名','機台','生產時間','訂單量','差異(生產量)','製令單號','批號']
        has_ship = False; has_dl = False
    colspan = 11 if has_ship else (10 if has_dl else 9)
    data_count = 0; rows_html = ''
    for _, r in df.iterrows():
        item = str(r['品項']) if pd.notna(r['品項']) else ''
        if not item or item == 'nan': continue
        if str(item).startswith('▌'):
            rows_html += f'<tr class="section-header"><td colspan="{colspan}">{esc(item)}</td></tr>\n'; continue
        if item == '品項': continue
        data_count += 1
        rows_html += (
            f'<tr>'
            f'<td>{esc(item)}</td>'
            f'<td style="font-size:12px">{esc(r["品號"])}</td>'
            f'<td style="font-size:12px">{esc(r["品名"])}</td>'
            f'<td>{esc(r["機台"])}</td><td>{esc(r["生產時間"])}</td>'
            + (f'<td>{esc(r.get("出貨日",""))}</td>' if has_ship else '')
            + (f'<td>{esc(r.get("交期",""))}</td>' if has_dl else '')
            + f'<td class="num">{fnum(r["訂單量"])}</td>'
            f'<td class="num">{fnum(r["差異(生產量)"])}</td>'
            f'<td style="font-size:12px">{clean_id(r["製令單號"])}</td>'
            f'<td style="font-size:12px">{clean_id(r["批號"])}</td>'
            f'</tr>\n'
        )
    ship_th = '<th>出貨日</th>' if has_ship else ''
    dl_th   = '<th>交期</th>'   if has_dl   else ''
    stats = (f'<div class="shortage-stats">'
             f'<span style="font-size:13px;color:#4a5568">共 <strong style="color:#1a56a0">{data_count}</strong> 筆須領料訂單</span>'
             f'</div>\n')
    table = (
        f'<div class="table-wrap" style="margin:10px 18px 24px">'
        f'<div class="table-scroll"><table><thead><tr>'
        f'<th>品項</th><th>品號</th><th>品名</th><th>機台</th><th>生產時間</th>'
        f'{ship_th}{dl_th}'
        f'<th class="num">訂單量</th><th class="num">差異</th><th>製令單號</th><th>批號</th>'
        f'</tr></thead><tbody>\n{rows_html}</tbody></table></div></div>'
    )
    return mascot_div('mascot-tab-pick', imgs_html) + src_div(fname, date_str) + stats + table


def gen_ready(path, fname, date_str, imgs_html):
    df = pd.read_excel(path, sheet_name='已備料', header=1)
    df.columns = ['訂單日期','序列','通路','品項','品號','品名',
                  '生產量','批號','製令單號','領料','入庫','交期','出貨日','備註']
    data_count = 0; rows_html = ''
    for _, r in df.iterrows():
        seq = r['序列']
        try:
            if pd.isna(seq): continue
        except: pass
        if str(seq) in ('nan','None',''): continue
        data_count += 1
        rows_html += (
            f'<tr>'
            f'<td>{ds(r["訂單日期"])}</td><td>{esc(r["序列"])}</td>'
            f'<td>{esc(r["通路"])}</td><td>{esc(r["品項"])}</td>'
            f'<td style="font-size:12px">{esc(r["品號"])}</td>'
            f'<td style="font-size:12px">{esc(r["品名"])}</td>'
            f'<td class="num">{fnum(r["生產量"])}</td>'
            f'<td style="font-size:12px">{clean_id(r["批號"])}</td>'
            f'<td style="font-size:12px">{clean_id(r["製令單號"])}</td>'
            f'<td>{ds(r["領料"])}</td><td>{ds(r["入庫"])}</td>'
            f'<td>{ds(r["交期"])}</td><td>{ds(r["出貨日"])}</td>'
            f'<td style="font-size:12px">{esc(r["備註"])}</td>'
            f'</tr>\n'
        )
    stats = (f'<div class="shortage-stats">'
             f'<span style="font-size:13px;color:#4a5568">共 <strong style="color:#1a56a0">{data_count}</strong> 筆已備料</span>'
             f'</div>\n')
    table = (
        f'<div class="table-wrap" style="margin:10px 18px 24px">'
        f'<div class="table-scroll"><table><thead><tr>'
        f'<th>訂單日期</th><th>序列</th><th>通路</th><th>品項</th>'
        f'<th>品號</th><th>品名</th><th class="num">生產量</th>'
        f'<th>批號</th><th>製令單號</th>'
        f'<th>領料</th><th>入庫</th><th>交期</th><th>出貨日</th><th>備註</th>'
        f'</tr></thead><tbody>\n{rows_html}</tbody></table></div></div>'
    )
    return mascot_div('mascot-tab-ready', imgs_html) + src_div(fname, date_str) + stats + table


# ════════════════════════════════════════════════════════
# 主程式
# ════════════════════════════════════════════════════════

# 1. 讀模板
if not os.path.exists(TMPL):
    print(f'[錯誤] 找不到模板：{TMPL}')
    sys.exit(1)
with open(TMPL, 'r', encoding='utf-8') as f:
    html = f.read()
print(f'讀取模板：{TMPL}')

# 2. 輪播圖片
imgs_html, nimgs = load_slideshow()

# 3. 進貨明細 Excel → RAW_DATA
inv_candidates = sorted(
    glob.glob(os.path.join(BASE,    '廠商進貨明細表*.xlsx')) +
    glob.glob(os.path.join(BASE,    '供應商進貨明細表*.xlsx')) +
    glob.glob(os.path.join(DESKTOP, '廠商進貨明細表*.xlsx')) +
    glob.glob(os.path.join(DESKTOP, '供應商進貨明細表*.xlsx')),
    key=os.path.getmtime
)
raw_json = '[]'; date_min = date_max = ''; inv_fname = ''
if not inv_candidates:
    print('[WARN] 找不到進貨明細 Excel')
else:
    inv_path = inv_candidates[-1]
    inv_fname = os.path.basename(inv_path)
    print(f'進貨明細：{inv_fname}')
    df = pd.read_excel(inv_path)
    df.columns = [re.sub(r'[\s　]+', '', c).strip() for c in df.columns]
    COL_MAP = {
        '廠商代號':'廠商代號','廠商簡稱':'廠商簡稱',
        '驗收/退貨日':'驗收日期','驗收日期':'驗收日期',
        '發票號碼':'發票號碼','品號':'品號','品名':'品名','規格':'規格','異動別':'異動別',
        '驗收/退貨單號':'驗收單號','進貨單號':'驗收單號',
        '驗收數量':'驗收數量','計價數量':'計價數量','單位':'單位',
        '原幣單位進價':'單價','單價':'單價',
        '本幣未稅金額':'本幣未稅金額','本幣稅額':'本幣稅額',
        '本幣金額合計':'本幣金額合計','採購單號':'採購單號','備註':'備註',
    }
    avail = {c: v for c, v in COL_MAP.items() if c in df.columns}
    df2 = df[list(avail.keys())].copy()
    df2.rename(columns=avail, inplace=True)
    df2['驗收日期'] = df2['驗收日期'].apply(lambda v: '' if pd.isna(v) else
        (pd.to_datetime(v).strftime('%Y/%m/%d') if not isinstance(v, str) else v))
    df2['發票號碼'] = df2['發票號碼'].fillna('').astype(str).str.strip().replace('nan','')
    for c in ['本幣未稅金額','本幣稅額','本幣金額合計','單價','驗收數量','計價數量']:
        if c in df2.columns:
            df2[c] = pd.to_numeric(df2[c], errors='coerce').fillna(0)
    for c in df2.columns:
        if df2[c].dtype == object:
            df2[c] = df2[c].fillna('').astype(str).str.strip().replace({'nan':'','None':''})
    dates = [d for d in df2['驗收日期'] if d]
    date_min = min(dates) if dates else ''
    date_max = max(dates) if dates else ''
    raw_json = json.dumps(df2.to_dict(orient='records'), ensure_ascii=False, separators=(',',':'))
    raw_json = re.sub(r':(?:NaN|-?Infinity)\b', ':null', raw_json)
    print(f'  → {len(df2)} 筆，{date_min} ~ {date_max}')

# 4. 原料品名對照 → MAT_NAME_DATA
mat_files = sorted(
    glob.glob(os.path.join(BASE,    '原料類品號*.xlsx')) +
    glob.glob(os.path.join(BASE,    '原料類品號*.XLSX')) +
    glob.glob(os.path.join(DESKTOP, '原料類品號*.xlsx')) +
    glob.glob(os.path.join(DESKTOP, '原料類品號*.XLSX')),
    key=os.path.getmtime
)
mat_map = {}
if mat_files:
    try:
        mdf = pd.read_excel(mat_files[-1], header=2)
        for _, r in mdf.iterrows():
            c = str(r.iloc[0]).strip(); n = str(r.iloc[1]).strip() if len(r) > 1 else ''
            if c and n and c not in ('nan','None') and n not in ('nan','None'):
                mat_map[c] = n
        print(f'原料品名：{len(mat_map)} 筆')
    except Exception as e:
        print(f'[WARN] 原料類品號讀取失敗：{e}')
mat_json = json.dumps(mat_map, ensure_ascii=False, separators=(',',':'))

# 5. 缺料分析 Excel → 各頁籤
shortage_folders = sorted(glob.glob(os.path.join(BASE, '缺料分析0*')))
shortage_path = None
if shortage_folders:
    latest = shortage_folders[-1]
    folder_tag = os.path.basename(latest)[4:]   # '0527'
    xlsx_list = sorted(glob.glob(os.path.join(latest, '缺料分析_*.xlsx')))
    if xlsx_list:
        shortage_path = xlsx_list[-1]
        shortage_date  = f'2026/{folder_tag[:2]}/{folder_tag[2:]}'
        shortage_fname = os.path.basename(shortage_path)
        print(f'缺料分析：{shortage_fname}（{shortage_date}）')

# ── 填入佔位符（全部 str.replace，不用 regex）────────────────
html = html.replace('__RAW_DATA_JSON__',       raw_json)
html = html.replace('__MAT_NAME_DATA_JSON__',  mat_json)
html = html.replace('__NIMGS__',               str(nimgs))
html = html.replace('__SLIDESHOW_IMGS__',      imgs_html)   # invoice tab mascot

if shortage_path:
    html = html.replace('__TAB_TAB_ORDERS__',    gen_orders(shortage_path, shortage_fname, shortage_date, imgs_html))
    html = html.replace('__TAB_TAB_SHORTAGE__',  gen_shortage(shortage_path, shortage_fname, shortage_date, imgs_html))
    html = html.replace('__TAB_TAB_INVENTORY__', gen_inventory(shortage_path, shortage_fname, shortage_date, imgs_html))
    html = html.replace('__TAB_TAB_MATERIAL__',  gen_material(shortage_path, shortage_fname, shortage_date, imgs_html))
    html = html.replace('__TAB_TAB_PICK__',      gen_pick(shortage_path, shortage_fname, shortage_date, imgs_html))
    html = html.replace('__TAB_TAB_READY__',     gen_ready(shortage_path, shortage_fname, shortage_date, imgs_html))
    print('[OK] 6 個頁籤已更新')
else:
    # 無缺料資料，清掉佔位符
    for p in ['__TAB_TAB_ORDERS__','__TAB_TAB_SHORTAGE__','__TAB_TAB_INVENTORY__',
              '__TAB_TAB_MATERIAL__','__TAB_TAB_PICK__','__TAB_TAB_READY__']:
        html = html.replace(p, '<div style="padding:20px;color:#888">（無缺料分析資料）</div>')
    print('[WARN] 找不到缺料分析，6個頁籤顯示空白')

# ── 確認沒有殘留佔位符 ────────────────────────────────────────
leftover = re.findall(r'__[A-Z_]+__', html)
if leftover:
    print(f'[WARN] 殘留佔位符：{sorted(set(leftover))}')
else:
    print('[OK] 所有佔位符已填入')

# ── 輸出 HTML ────────────────────────────────────────────────
for out_path in OUTPUTS:
    with open(out_path, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f'[OK] 儲存：{out_path}')

print('\n完成！')
