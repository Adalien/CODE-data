# -*- coding: utf-8 -*-
"""
產生廠商進貨明細查詢 HTML
- 自動讀取最新的進貨明細 Excel 和最新缺料分析 Excel
- 更新現有 HTML（保留 CSS / JS / 圖片不變，只換資料）
- 同時更新桌面版與 CODE資料版
"""
import pandas as pd
import json
import re
import glob
import os
import sys
from datetime import datetime

# ── 路徑設定 ──────────────────────────────────────────────────
BASE    = r'C:\Users\admin\OneDrive\桌面\CODE資料'
DESKTOP = r'C:\Users\admin\OneDrive\桌面'
HTML_DST  = os.path.join(DESKTOP, '進貨明細查詢.html')
HTML_REPO = os.path.join(BASE,    '進貨明細查詢.html')

# ── 工具函式 ───────────────────────────────────────────────────
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
    """去掉製令單號/批號尾端多餘的 .0（Excel 數字轉字串問題）"""
    if v is None: return ''
    try:
        if pd.isna(v): return ''
    except: pass
    s = str(v).strip()
    if s in ('nan', 'None', ''): return ''
    if s.endswith('.0') and s[:-2].replace('-','').isdigit(): return s[:-2]
    return s

def replace_tab(html, tab_id, new_inner):
    pat = rf'(<div id="{tab_id}"[^>]*>)(.*?)(</div><!--\s*/{tab_id}\s*-->)'
    return re.sub(pat, rf'\g<1>{new_inner}\g<3>', html, flags=re.DOTALL)

def data_source_div(fname, date_str):
    return f'<div class="data-source">資料來源：{esc(fname)}（{date_str}）</div>\n'

# ══════════════════════════════════════════════════════════════
# 一、讀取進貨明細 Excel
# ══════════════════════════════════════════════════════════════
inv_candidates = (
    sorted(glob.glob(os.path.join(BASE,    '廠商進貨明細表*.xlsx'))) +
    sorted(glob.glob(os.path.join(DESKTOP, '廠商進貨明細表*.xlsx')))
)
if not inv_candidates:
    print('[警告] 找不到廠商進貨明細表 Excel，進貨明細頁籤維持現有資料')
    inv_path = None
else:
    inv_path = inv_candidates[-1]
    print(f'進貨明細: {os.path.basename(inv_path)}')

raw_json = None
date_min = date_max = ''
inv_fname = ''
if inv_path:
    df = pd.read_excel(inv_path)
    df.columns = [re.sub(r'[\s　]+', '', c).strip() for c in df.columns]
    COL_MAP = {
        '廠商代號':'廠商代號','廠商簡稱':'廠商簡稱',
        '驗收/退貨日':'驗收日期','發票號碼':'發票號碼',
        '品號':'品號','品名':'品名','規格':'規格','異動別':'異動別',
        '驗收/退貨單號':'驗收單號','驗收數量':'驗收數量',
        '計價數量':'計價數量','單位':'單位','單價':'單價',
        '本幣未稅金額':'本幣未稅金額','本幣稅額':'本幣稅額',
        '本幣金額合計':'本幣金額合計','採購單號':'採購單號','備註':'備註',
    }
    avail = {c: v for c, v in COL_MAP.items() if c in df.columns}
    df2 = df[list(avail.keys())].copy()
    df2.rename(columns=avail, inplace=True)

    def to_date_str(v):
        if pd.isna(v): return ''
        try: return pd.to_datetime(v).strftime('%Y/%m/%d')
        except: return str(v)

    df2['驗收日期'] = df2['驗收日期'].apply(to_date_str)
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
    inv_fname = os.path.basename(inv_path)
    print(f'  → {len(df2)} 筆，{date_min} ~ {date_max}')

# ══════════════════════════════════════════════════════════════
# 二、找最新缺料分析 Excel
# ══════════════════════════════════════════════════════════════
shortage_folders = sorted(glob.glob(os.path.join(BASE, '缺料分析0*')))
if not shortage_folders:
    print('[警告] 找不到缺料分析資料夾，缺料頁籤維持現有資料')
    shortage_path = None
else:
    latest_folder = shortage_folders[-1]
    folder_tag = os.path.basename(latest_folder)[4:]   # '0527'
    # 優先最高版次（_5 > _4 > 無後綴）
    xlsx_cands = sorted(glob.glob(os.path.join(latest_folder, '缺料分析_*.xlsx')))
    shortage_path = xlsx_cands[-1] if xlsx_cands else None
    if shortage_path:
        shortage_date = f'2026/{folder_tag[:2]}/{folder_tag[2:]}'
        shortage_fname = os.path.basename(shortage_path)
        # 若檔案被 Excel 鎖定，用 Windows robocopy 建立暫存複本
        try:
            with open(shortage_path, 'rb') as _f:
                _f.read(4)
        except PermissionError:
            import tempfile, subprocess as _sp
            _tmp_dir = tempfile.gettempdir()
            _sp.run(['robocopy', os.path.dirname(shortage_path),
                     _tmp_dir, shortage_fname, '/NFL', '/NDL', '/NJH', '/NJS'],
                    capture_output=True)
            _tmp = os.path.join(_tmp_dir, shortage_fname)
            if os.path.exists(_tmp):
                shortage_path = _tmp
                print('  [警告] 原檔被鎖定，改用暫存複本')
            else:
                print('  [警告] 缺料分析檔被Excel鎖定且無法複製，略過缺料頁籤')
                shortage_path = None
        print(f'缺料分析: {shortage_fname}（{shortage_date}）')

# ══════════════════════════════════════════════════════════════
# 三、缺料分析各頁籤 HTML 生成函式
# ══════════════════════════════════════════════════════════════

def gen_shortage(path, fname, date_str):
    """缺料清單"""
    df = pd.read_excel(path, sheet_name='缺料清單', header=1)
    # 欄位數依版本不同（13欄含最早交期，12欄舊版）
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
    return data_source_div(fname, date_str) + stats + table


def gen_inventory(path, fname, date_str):
    """庫存現況"""
    df = pd.read_excel(path, sheet_name='庫存解析驗證', header=1)
    df.columns = ['類別','庫存原始字串','解析量','單位','備註']
    rows_html = ''
    for _, r in df.iterrows():
        cat = str(r['類別']) if pd.notna(r['類別']) else ''
        orig = str(r['庫存原始字串']) if pd.notna(r['庫存原始字串']) else ''
        if not cat or cat == 'nan': continue
        if str(cat).startswith('【'):
            rows_html += f'<tr class="cat-header"><td colspan="5">{esc(cat)}</td></tr>\n'
            continue
        qty = r['解析量']
        qty_str = fnum(qty)
        unit = str(r['單位']) if pd.notna(r['單位']) else ''
        note = str(r['備註']) if pd.notna(r['備註']) else ''
        is_zero = False
        try: is_zero = float(qty) == 0
        except: pass
        cls = 'inv-zero' if is_zero else ''
        rows_html += (
            f'<tr class="{cls}">'
            f'<td>{esc(cat)}</td>'
            f'<td style="font-size:13px">{esc(orig)}</td>'
            f'<td class="num">{qty_str}</td>'
            f'<td>{esc(unit)}</td>'
            f'<td style="font-size:12px">{esc(note)}</td>'
            f'</tr>\n'
        )
    table = (
        f'<div class="table-wrap" style="margin:10px 18px 24px">'
        f'<div class="table-scroll"><table>\n'
        f'<thead><tr><th>類別</th><th>庫存原始字串</th>'
        f'<th class="num">解析量</th><th>單位</th><th>備註</th></tr></thead>'
        f'<tbody>\n{rows_html}</tbody></table></div></div>'
    )
    return data_source_div(fname, date_str) + table


def gen_material(path, fname, date_str):
    """材料需求彙總"""
    df = pd.read_excel(path, sheet_name='材料需求彙總', header=0)
    # 正規化欄位名稱
    df.columns = [str(c).strip() for c in df.columns]
    lack_n   = (df['狀態'].str.contains('缺料', na=False)).sum()
    ok_n     = len(df) - lack_n
    stats = (
        f'<div class="shortage-stats">\n'
        f'<span style="font-size:13px;color:#4a5568">總品項 <strong style="color:#1a56a0">{len(df)}</strong></span>\n'
        f'<span class="stat-badge stat-red">⚠ 缺料 {lack_n}</span>\n'
        f'<span class="stat-badge" style="background:#f0fff4;color:#276749">✓ 充足 {ok_n}</span>\n'
        f'</div>\n'
    )
    unit_col  = '單位'
    unit1_col = '單位.1' if '單位.1' in df.columns else '單位'
    rows_html = ''
    for _, r in df.iterrows():
        status = str(r['狀態']) if pd.notna(r.get('狀態')) else ''
        cls = 'lack-severe' if '缺料' in status else 'inv-ok'
        rows_html += (
            f'<tr class="{cls}">'
            f'<td>{esc(r["材料類型"])}</td>'
            f'<td>{esc(r["材料品號"])}</td>'
            f'<td style="font-size:12px">{esc(r.get("品名(系統)",""))}</td>'
            f'<td>{esc(r.get("材料品名","") if pd.notna(r.get("材料品名")) else "")}</td>'
            f'<td class="num">{fnum(r.get("總需求量",""))}</td>'
            f'<td>{esc(r.get(unit_col,""))}</td>'
            f'<td class="num">{fnum(r.get("庫存量",""))}</td>'
            f'<td>{esc(r.get(unit1_col,""))}</td>'
            f'<td class="num">{fnum(r.get("淨缺量",""))}</td>'
            f'<td>{esc(status)}</td>'
            f'</tr>\n'
        )
    table = (
        f'<div class="table-wrap" style="margin:10px 18px 24px">'
        f'<div class="table-scroll"><table><thead><tr>'
        f'<th>材料類型</th><th>材料品號</th><th style="font-size:12px">品名(系統)</th>'
        f'<th>材料品名</th>'
        f'<th class="num">總需求量</th><th>單位</th>'
        f'<th class="num">庫存量</th><th>單位</th>'
        f'<th class="num">淨缺量</th><th>狀態</th>'
        f'</tr></thead><tbody>\n{rows_html}</tbody></table></div></div>'
    )
    return data_source_div(fname, date_str) + stats + table


def gen_pick(path, fname, date_str):
    """須領料"""
    df = pd.read_excel(path, sheet_name='須領料', header=0)
    # 欄位數依版本：11欄(含出貨日+交期) / 10欄(含交期) / 9欄(舊版)
    ncols = len(df.columns)
    if ncols >= 11:
        df.columns = ['品項','品號','品名','機台','生產時間','出貨日','交期','訂單量','差異(生產量)','製令單號','批號']
        has_ship = True; has_deadline = True
    elif ncols >= 10:
        df.columns = ['品項','品號','品名','機台','生產時間','交期','訂單量','差異(生產量)','製令單號','批號']
        has_ship = False; has_deadline = True
    else:
        df.columns = ['品項','品號','品名','機台','生產時間','訂單量','差異(生產量)','製令單號','批號']
        has_ship = False; has_deadline = False
    data_count = 0
    rows_html  = ''
    colspan = 11 if has_ship else (10 if has_deadline else 9)
    for _, r in df.iterrows():
        item = str(r['品項']) if pd.notna(r['品項']) else ''
        if not item or item == 'nan': continue
        if str(item).startswith('▌'):
            rows_html += f'<tr class="section-header"><td colspan="{colspan}">{esc(item)}</td></tr>\n'
            continue
        if item == '品項': continue
        data_count += 1
        ship_td = f'<td>{esc(r.get("出貨日",""))}</td>' if has_ship else ''
        dl_td   = f'<td>{esc(r.get("交期",""))}</td>'   if has_deadline else ''
        rows_html += (
            f'<tr>'
            f'<td>{esc(item)}</td>'
            f'<td style="font-size:12px">{esc(r["品號"])}</td>'
            f'<td style="font-size:12px">{esc(r["品名"])}</td>'
            f'<td>{esc(r["機台"])}</td>'
            f'<td>{esc(r["生產時間"])}</td>'
            f'{ship_td}'
            f'{dl_td}'
            f'<td class="num">{fnum(r["訂單量"])}</td>'
            f'<td class="num">{fnum(r["差異(生產量)"])}</td>'
            f'<td style="font-size:12px">{clean_id(r["製令單號"])}</td>'
            f'<td style="font-size:12px">{clean_id(r["批號"])}</td>'
            f'</tr>\n'
        )
    stats = (
        f'<div class="shortage-stats">'
        f'<span style="font-size:13px;color:#4a5568">共 '
        f'<strong style="color:#1a56a0">{data_count}</strong> 筆須領料訂單</span>'
        f'</div>\n'
    )
    ship_th = '<th>出貨日</th>' if has_ship else ''
    dl_th   = '<th>交期</th>'   if has_deadline else ''
    table = (
        f'<div class="table-wrap" style="margin:10px 18px 24px">'
        f'<div class="table-scroll"><table><thead><tr>'
        f'<th>品項</th><th>品號</th><th>品名</th><th>機台</th><th>生產時間</th>'
        f'{ship_th}{dl_th}'
        f'<th class="num">訂單量</th><th class="num">差異</th><th>製令單號</th><th>批號</th>'
        f'</tr></thead><tbody>\n{rows_html}</tbody></table></div></div>'
    )
    return data_source_div(fname, date_str) + stats + table


def gen_ready(path, fname, date_str):
    """已備料"""
    df = pd.read_excel(path, sheet_name='已備料', header=1)
    df.columns = ['訂單日期','序列','通路','品項','品號','品名',
                  '生產量','批號','製令單號','領料','入庫','交期','出貨日','備註']
    data_count = 0
    rows_html  = ''
    for _, r in df.iterrows():
        seq = r['序列']
        try:
            if pd.isna(seq): continue
        except: pass
        try:
            if str(seq) in ('nan','None',''): continue
        except: pass
        data_count += 1
        def ds(v):
            if pd.isna(v): return ''
            try: return pd.to_datetime(v).strftime('%Y/%m/%d')
            except: return str(v)
        rows_html += (
            f'<tr>'
            f'<td>{ds(r["訂單日期"])}</td>'
            f'<td>{esc(r["序列"])}</td>'
            f'<td>{esc(r["通路"])}</td>'
            f'<td>{esc(r["品項"])}</td>'
            f'<td style="font-size:12px">{esc(r["品號"])}</td>'
            f'<td style="font-size:12px">{esc(r["品名"])}</td>'
            f'<td class="num">{fnum(r["生產量"])}</td>'
            f'<td style="font-size:12px">{clean_id(r["批號"])}</td>'
            f'<td style="font-size:12px">{clean_id(r["製令單號"])}</td>'
            f'<td>{ds(r["領料"])}</td>'
            f'<td>{ds(r["入庫"])}</td>'
            f'<td>{ds(r["交期"])}</td>'
            f'<td>{ds(r["出貨日"])}</td>'
            f'<td style="font-size:12px">{esc(r["備註"])}</td>'
            f'</tr>\n'
        )
    stats = (
        f'<div class="shortage-stats">'
        f'<span style="font-size:13px;color:#4a5568">共 '
        f'<strong style="color:#1a56a0">{data_count}</strong> 筆已備料</span>'
        f'</div>\n'
    )
    table = (
        f'<div class="table-wrap" style="margin:10px 18px 24px">'
        f'<div class="table-scroll"><table><thead><tr>'
        f'<th>訂單日期</th><th>序列</th><th>通路</th><th>品項</th>'
        f'<th>品號</th><th>品名</th><th class="num">生產量</th>'
        f'<th>批號</th><th>製令單號</th>'
        f'<th>領料</th><th>入庫</th><th>交期</th><th>出貨日</th><th>備註</th>'
        f'</tr></thead><tbody>\n{rows_html}</tbody></table></div></div>'
    )
    return data_source_div(fname, date_str) + stats + table

def gen_orders(path, fname, date_str):
    """訂單總表（訂單交期總覽工作表）"""
    try:
        df = pd.read_excel(path, sheet_name='訂單交期總覽', header=1)
    except Exception:
        return '<div style="padding:20px;color:#888">找不到訂單交期總覽工作表</div>'
    df.columns = [str(c).strip() for c in df.columns]
    def ds(v):
        if pd.isna(v): return ''
        try: return pd.to_datetime(v).strftime('%Y/%m/%d')
        except:
            s = str(v).strip()
            return '' if s in ('nan','NaT','None','') else s
    # 今日日期判斷（交期7天內標色）
    from datetime import date
    today = pd.Timestamp(date.today())
    rows_html = ''
    data_count = 0
    for _, r in df.iterrows():
        pno = str(r.get('品號','')).strip()
        if not pno or pno in ('nan','None',''): continue
        data_count += 1
        dl_raw = r.get('交期','')
        try:
            dl_ts = pd.to_datetime(dl_raw) if dl_raw and str(dl_raw) not in ('nan','NaT','') else None
        except: dl_ts = None
        if dl_ts and (dl_ts - today).days <= 7:
            cls = 'lack-severe'
        else:
            cls = ''
        rows_html += (
            f'<tr class="{cls}">'
            f'<td>{ds(r.get("訂單日期",""))}</td>'
            f'<td>{esc(r.get("序列",""))}</td>'
            f'<td>{esc(r.get("通路",""))}</td>'
            f'<td>{esc(r.get("品項",""))}</td>'
            f'<td style="font-size:12px">{esc(r.get("品號",""))}</td>'
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
        f'</tr></thead><tbody>\n{rows_html}</tbody></table></div></div>'
    )
    return data_source_div(fname, date_str) + stats + table

# ══════════════════════════════════════════════════════════════
# 四、讀現有 HTML 並更新
# ══════════════════════════════════════════════════════════════
template_path = HTML_REPO if os.path.exists(HTML_REPO) else HTML_DST
print(f'\n讀取 HTML 模板: {template_path}')
with open(template_path, 'r', encoding='utf-8') as f:
    html = f.read()

# ── 把 rabbit-zone 移到全域固定位置（所有 tab 都顯示）──────────
if '<div class="rabbit-zone">' in html:
    # 1. 取出 rabbit-zone HTML
    _rz_start = html.find('<div class="rabbit-zone">')
    _rz_end   = html.find('</div>', _rz_start) + 6
    _rz_html  = html[_rz_start:_rz_end]

    # 2. 從原本的位置移除（進貨明細 filter-card 裡）
    html = html[:_rz_start] + html[_rz_end:]

    # 3. 確保 CSS 為 fixed 定位
    _fixed_css = '''
  /* ── 賤兔固定在右下角（全頁可見）── */
  .rabbit-zone-fixed {
    position: fixed;
    bottom: 18px;
    right: 18px;
    width: 160px;
    height: 100px;
    z-index: 999;
    pointer-events: none;
  }
  .rabbit-zone-fixed .rabbit-slide {
    position: absolute;
    bottom: 0; left: 0;
    width: 160px; height: 100px;
    object-fit: contain;
    opacity: 0;
    transition: opacity 0.6s ease-in-out;
  }
  .rabbit-zone-fixed .rabbit-slide.active { opacity: 1; }
'''
    if 'rabbit-zone-fixed' not in html:
        html = html.replace('</style>', _fixed_css + '</style>', 1)

    # 4. 把 rabbit-zone 加回 body 底部（fixed 定位，全頁顯示）
    _rz_fixed = _rz_html.replace(
        '<div class="rabbit-zone">',
        '<div class="rabbit-zone rabbit-zone-fixed">'
    )
    html = html.replace('</body>', _rz_fixed + '\n</body>', 1)
    print('賤兔輪播已移至全頁固定位置（右下角）')

# ── 更新缺料分析五個頁籤 ──────────────────────────────────────
if shortage_path:
    # ── 新增「訂單總表」頁籤按鈕（若尚未存在）──
    if 'tab-orders' not in html:
        html = html.replace(
            '<button class="tab-btn" onclick="switchTab(\'tab-shortage\', this)">🔴 缺料清單</button>',
            '<button class="tab-btn" onclick="switchTab(\'tab-orders\', this)">📅 訂單總表</button>\n<button class="tab-btn" onclick="switchTab(\'tab-shortage\', this)">🔴 缺料清單</button>'
        )
        html = html.replace(
            '<div id="tab-shortage" class="tab-content">',
            '<div id="tab-orders" class="tab-content"></div><!--/tab-orders-->\n<div id="tab-shortage" class="tab-content">'
        )
    print('更新訂單總表頁籤...')
    html = replace_tab(html, 'tab-orders',    gen_orders(shortage_path, shortage_fname, shortage_date))
    print('更新缺料清單頁籤...')
    html = replace_tab(html, 'tab-shortage',  gen_shortage(shortage_path, shortage_fname, shortage_date))
    print('更新庫存現況頁籤...')
    html = replace_tab(html, 'tab-inventory', gen_inventory(shortage_path, shortage_fname, shortage_date))
    print('更新材料需求頁籤...')
    html = replace_tab(html, 'tab-material',  gen_material(shortage_path, shortage_fname, shortage_date))
    print('更新須領料頁籤...')
    html = replace_tab(html, 'tab-pick',      gen_pick(shortage_path, shortage_fname, shortage_date))
    print('更新已備料頁籤...')
    html = replace_tab(html, 'tab-ready',     gen_ready(shortage_path, shortage_fname, shortage_date))

# ── 更新進貨明細資料（RAW_DATA）──────────────────────────────
if raw_json:
    print('更新進貨明細 RAW_DATA...')
    html = re.sub(
        r'const RAW_DATA = \[.*?\];',
        f'const RAW_DATA = {raw_json};',
        html, flags=re.DOTALL
    )
    # 更新頁首說明文字
    html = re.sub(
        r'<p>資料範圍：[^<]+</p>',
        f'<p>資料範圍：{date_min} ～ {date_max}　共 {len(df2)} 筆　(來源：{esc(inv_fname)})</p>',
        html
    )

# ── 存檔（桌面 + CODE資料 + 分享版）─────────────────────────
SHARE_DST  = os.path.join(DESKTOP, '達特世生技_進貨明細查詢_2026.html')
SHARE_REPO = os.path.join(BASE,    '達特世生技_進貨明細查詢_2026.html')
for out in [HTML_DST, HTML_REPO, SHARE_DST, SHARE_REPO]:
    with open(out, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f'[OK] 已存: {out}')

print('\n完成！')
