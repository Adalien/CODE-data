# -*- coding: utf-8 -*-
"""
產生廠商進貨明細查詢 HTML（獨立單頁，雙擊可開啟）
"""
import pandas as pd
import json
import re
from datetime import datetime

# ── 讀取 Excel ──────────────────────────────────────────────
df = pd.read_excel(r'C:\Users\admin\OneDrive\桌面\廠商進貨明細表2025-2026-5-26.xlsx')

# 欄位名稱：去除全形／半形空格
df.columns = [re.sub(r'[\s　]+', '', c).strip() for c in df.columns]

# 實際欄位對應（清理後）
COL_MAP = {
    '廠商代號':    '廠商代號',
    '廠商簡稱':    '廠商簡稱',
    '驗收/退貨日': '驗收日期',
    '發票號碼':    '發票號碼',
    '品號':        '品號',
    '品名':        '品名',
    '規格':        '規格',
    '異動別':      '異動別',
    '驗收/退貨單號': '驗收單號',
    '驗收數量':    '驗收數量',
    '計價數量':    '計價數量',
    '單位':        '單位',
    '單價':        '單價',
    '本幣未稅金額': '本幣未稅金額',
    '本幣稅額':    '本幣稅額',
    '本幣金額合計': '本幣金額合計',
    '採購單號':    '採購單號',
    '備註':        '備註',
}

# 只保留需要的欄位（依實際存在的欄位過濾）
available = {c: v for c, v in COL_MAP.items() if c in df.columns}
df2 = df[list(available.keys())].copy()
df2.rename(columns=available, inplace=True)

# ── 日期欄轉字串 ──────────────────────────────────────────
def to_date_str(v):
    if pd.isna(v):
        return ''
    try:
        return pd.to_datetime(v).strftime('%Y/%m/%d')
    except:
        return str(v)

df2['驗收日期'] = df2['驗收日期'].apply(to_date_str)
df2['發票號碼'] = df2['發票號碼'].fillna('').astype(str).str.strip()
df2['發票號碼'] = df2['發票號碼'].replace('nan', '')

# ── 金額欄位轉數字 ─────────────────────────────────────────
money_cols = ['本幣未稅金額', '本幣稅額', '本幣金額合計', '單價', '驗收數量', '計價數量']
for c in money_cols:
    if c in df2.columns:
        df2[c] = pd.to_numeric(df2[c], errors='coerce').fillna(0)

# ── 其他文字欄 NaN → '' ────────────────────────────────────
for c in df2.columns:
    if df2[c].dtype == object:
        df2[c] = df2[c].fillna('').astype(str).str.strip()
        df2[c] = df2[c].replace('nan', '').replace('None', '')

# ── 資料範圍日期 ───────────────────────────────────────────
dates = [d for d in df2['驗收日期'] if d]
date_min = min(dates) if dates else ''
date_max = max(dates) if dates else ''

# ── 轉 JSON ───────────────────────────────────────────────
records = df2.to_dict(orient='records')
json_data = json.dumps(records, ensure_ascii=False, separators=(',', ':'))

print(f'資料筆數: {len(records)}')
print(f'日期範圍: {date_min} ~ {date_max}')

# ── 產生 HTML ──────────────────────────────────────────────
html = f"""<!DOCTYPE html>
<html lang="zh-TW">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>達特世生技 廠商進貨明細查詢</title>
<style>
  *, *::before, *::after {{ box-sizing: border-box; }}
  body {{
    font-family: '微軟正黑體', 'Microsoft JhengHei', 'Noto Sans TC', sans-serif;
    background: #f0f4f8;
    margin: 0;
    padding: 0;
    color: #2d3748;
    font-size: 14px;
  }}
  /* ── 頂部 header ── */
  .header {{
    background: linear-gradient(135deg, #1a56a0 0%, #2563eb 100%);
    color: #fff;
    padding: 18px 24px 14px;
    box-shadow: 0 2px 8px rgba(0,0,0,0.18);
  }}
  .header h1 {{ margin: 0; font-size: 22px; font-weight: 700; letter-spacing: 1px; }}
  .header p  {{ margin: 4px 0 0; font-size: 13px; opacity: 0.85; }}

  /* ── 篩選區 ── */
  .filter-card {{
    background: #fff;
    margin: 18px 18px 0;
    border-radius: 10px;
    box-shadow: 0 1px 4px rgba(0,0,0,0.10);
    padding: 16px 20px 12px;
  }}
  .filter-row {{
    display: flex;
    flex-wrap: wrap;
    gap: 10px 16px;
    align-items: flex-end;
  }}
  .filter-group {{
    display: flex;
    flex-direction: column;
    gap: 4px;
    min-width: 140px;
  }}
  .filter-group label {{
    font-size: 12px;
    font-weight: 600;
    color: #4a5568;
  }}
  .filter-group input,
  .filter-group select {{
    height: 34px;
    padding: 0 10px;
    border: 1px solid #cbd5e0;
    border-radius: 6px;
    font-size: 13px;
    font-family: inherit;
    color: #2d3748;
    background: #f7fafc;
    outline: none;
    transition: border-color .15s;
  }}
  .filter-group input:focus,
  .filter-group select:focus {{
    border-color: #2563eb;
    background: #fff;
  }}
  .filter-group.wide input {{ width: 220px; }}
  .btn {{
    height: 34px;
    padding: 0 20px;
    border: none;
    border-radius: 6px;
    font-size: 13px;
    font-family: inherit;
    cursor: pointer;
    font-weight: 600;
    transition: background .15s, transform .1s;
  }}
  .btn:active {{ transform: scale(0.97); }}
  .btn-primary {{ background: #2563eb; color: #fff; }}
  .btn-primary:hover {{ background: #1d4ed8; }}
  .btn-secondary {{ background: #e2e8f0; color: #4a5568; }}
  .btn-secondary:hover {{ background: #cbd5e0; }}
  .btn-row {{
    display: flex;
    gap: 8px;
    align-items: flex-end;
    padding-bottom: 0;
  }}

  /* ── 統計摘要 ── */
  .summary-bar {{
    background: #fff;
    margin: 10px 18px 0;
    border-radius: 10px;
    box-shadow: 0 1px 4px rgba(0,0,0,0.08);
    padding: 10px 20px;
    display: flex;
    flex-wrap: wrap;
    gap: 6px 24px;
    align-items: center;
  }}
  .summary-item {{
    font-size: 13px;
    color: #4a5568;
  }}
  .summary-item span {{
    font-weight: 700;
    color: #1a56a0;
    font-size: 15px;
  }}
  .summary-label {{ margin-right: 4px; }}

  /* ── 表格容器 ── */
  .table-wrap {{
    margin: 10px 18px 24px;
    background: #fff;
    border-radius: 10px;
    box-shadow: 0 1px 4px rgba(0,0,0,0.08);
    overflow: hidden;
  }}
  .table-scroll {{
    overflow-x: auto;
    max-height: calc(100vh - 300px);
  }}
  table {{
    width: 100%;
    border-collapse: collapse;
    font-size: 13px;
    min-width: 1100px;
  }}
  thead tr {{
    background: #1a56a0;
    color: #fff;
    position: sticky;
    top: 0;
    z-index: 2;
  }}
  thead th {{
    padding: 10px 10px;
    text-align: left;
    white-space: nowrap;
    font-weight: 600;
    cursor: pointer;
    user-select: none;
    border-right: 1px solid rgba(255,255,255,0.12);
    transition: background .1s;
  }}
  thead th:hover {{ background: #1d4ed8; }}
  thead th.sort-asc::after  {{ content: ' ▲'; font-size: 10px; opacity: .8; }}
  thead th.sort-desc::after {{ content: ' ▼'; font-size: 10px; opacity: .8; }}
  thead th.num  {{ text-align: right; }}

  tbody tr:nth-child(even) {{ background: #f7fafc; }}
  tbody tr:hover {{ background: #ebf4ff; }}
  tbody td {{
    padding: 7px 10px;
    border-bottom: 1px solid #e8edf2;
    white-space: nowrap;
    border-right: 1px solid #e8edf2;
  }}
  tbody td.num {{ text-align: right; font-variant-numeric: tabular-nums; }}

  /* 廠商 badge */
  .badge {{
    display: inline-block;
    padding: 2px 8px;
    border-radius: 20px;
    font-size: 11px;
    font-weight: 600;
    white-space: nowrap;
    color: #fff;
  }}

  /* ── 分頁 ── */
  .pagination {{
    display: flex;
    justify-content: center;
    align-items: center;
    gap: 6px;
    padding: 12px 16px;
    border-top: 1px solid #e8edf2;
    flex-wrap: wrap;
  }}
  .page-btn {{
    min-width: 32px;
    height: 30px;
    padding: 0 8px;
    border: 1px solid #cbd5e0;
    border-radius: 5px;
    background: #fff;
    color: #4a5568;
    font-size: 13px;
    font-family: inherit;
    cursor: pointer;
    transition: all .1s;
  }}
  .page-btn:hover {{ background: #ebf4ff; border-color: #2563eb; color: #2563eb; }}
  .page-btn.active {{ background: #2563eb; border-color: #2563eb; color: #fff; font-weight: 700; }}
  .page-btn:disabled {{ opacity: .35; cursor: default; }}
  .page-info {{ font-size: 13px; color: #718096; margin: 0 6px; }}

  /* ── RWD ── */
  @media (max-width: 768px) {{
    .filter-card {{ margin: 10px 8px 0; padding: 12px 12px 8px; }}
    .table-wrap  {{ margin: 8px 8px 16px; }}
    .summary-bar {{ margin: 8px 8px 0; padding: 8px 12px; }}
    .header {{ padding: 12px 14px 10px; }}
    .header h1 {{ font-size: 17px; }}
    .filter-group.wide input {{ width: 100%; }}
  }}
  .no-data {{
    text-align: center;
    padding: 40px;
    color: #a0aec0;
    font-size: 15px;
  }}
</style>
</head>
<body>

<div class="header">
  <h1>達特世生技 廠商進貨明細查詢</h1>
  <p>資料範圍：{date_min} ～ {date_max}　共 {len(records)} 筆　(來源：廠商進貨明細表2025-2026-5-26.xlsx)</p>
</div>

<div class="filter-card">
  <div class="filter-row">
    <div class="filter-group">
      <label>廠商</label>
      <select id="selVendor">
        <option value="">全部廠商</option>
      </select>
    </div>
    <div class="filter-group">
      <label>驗收日期（起）</label>
      <input type="date" id="dateFrom">
    </div>
    <div class="filter-group">
      <label>驗收日期（訖）</label>
      <input type="date" id="dateTo">
    </div>
    <div class="filter-group wide">
      <label>品號 / 品名（模糊搜尋）</label>
      <input type="text" id="txtItem" placeholder="輸入品號或品名...">
    </div>
    <div class="filter-group">
      <label>採購單號</label>
      <input type="text" id="txtPO" placeholder="採購單號...">
    </div>
    <div class="btn-row">
      <button class="btn btn-primary" onclick="doSearch()">查詢</button>
      <button class="btn btn-secondary" onclick="doClear()">清除</button>
    </div>
  </div>
</div>

<div class="summary-bar" id="summaryBar">
  <div class="summary-item"><span class="summary-label">符合筆數</span><span id="sumCount">0</span></div>
  <div class="summary-item"><span class="summary-label">本幣未稅合計</span>NT$ <span id="sumUntax">0</span></div>
  <div class="summary-item"><span class="summary-label">稅額合計</span>NT$ <span id="sumTax">0</span></div>
  <div class="summary-item"><span class="summary-label">本幣金額合計</span>NT$ <span id="sumTotal">0</span></div>
</div>

<div class="table-wrap">
  <div class="table-scroll">
    <table id="dataTable">
      <thead>
        <tr>
          <th data-col="驗收日期">驗收日期</th>
          <th data-col="廠商簡稱">廠商簡稱</th>
          <th data-col="品號">品號</th>
          <th data-col="品名">品名</th>
          <th data-col="規格">規格</th>
          <th data-col="驗收數量" class="num">驗收數量</th>
          <th data-col="單位">單位</th>
          <th data-col="單價" class="num">單價</th>
          <th data-col="本幣未稅金額" class="num">本幣未稅金額</th>
          <th data-col="本幣稅額" class="num">本幣稅額</th>
          <th data-col="本幣金額合計" class="num">本幣金額合計</th>
          <th data-col="採購單號">採購單號</th>
          <th data-col="發票號碼">發票號碼</th>
          <th data-col="備註">備註</th>
        </tr>
      </thead>
      <tbody id="tbody"></tbody>
    </table>
  </div>
  <div class="pagination" id="pagination"></div>
</div>

<script>
// ── 內嵌資料 ──────────────────────────────────────────────
const RAW_DATA = {json_data};

// ── 廠商顏色 badge ─────────────────────────────────────────
const BADGE_COLORS = [
  '#2563eb','#16a34a','#dc2626','#9333ea','#d97706',
  '#0891b2','#be185d','#7c3aed','#c2410c','#0f766e',
  '#1e40af','#15803d','#b91c1c','#6d28d9','#b45309',
];
const vendorColors = {{}};
let colorIdx = 0;
function getVendorColor(name) {{
  if (!vendorColors[name]) {{
    vendorColors[name] = BADGE_COLORS[colorIdx % BADGE_COLORS.length];
    colorIdx++;
  }}
  return vendorColors[name];
}}

// ── 狀態 ───────────────────────────────────────────────────
let filteredData = [...RAW_DATA];
let currentPage  = 1;
const PAGE_SIZE  = 50;
let sortCol      = '驗收日期';
let sortAsc      = true;

// ── 初始化 ─────────────────────────────────────────────────
function init() {{
  // 廠商下拉
  const vendors = [...new Set(RAW_DATA.map(r => r['廠商簡稱']).filter(v => v))].sort();
  const sel = document.getElementById('selVendor');
  vendors.forEach(v => {{
    const opt = document.createElement('option');
    opt.value = v; opt.textContent = v;
    sel.appendChild(opt);
  }});
  // 日期預設
  const dates = RAW_DATA.map(r => r['驗收日期']).filter(d => d).sort();
  if (dates.length) {{
    document.getElementById('dateFrom').value = toInputDate(dates[0]);
    document.getElementById('dateTo').value   = toInputDate(dates[dates.length - 1]);
  }}
  // 排序表頭
  document.querySelectorAll('thead th[data-col]').forEach(th => {{
    th.addEventListener('click', () => {{
      const col = th.dataset.col;
      if (sortCol === col) {{ sortAsc = !sortAsc; }}
      else {{ sortCol = col; sortAsc = true; }}
      applySort();
      renderTable();
      updateSortHeaders();
    }});
  }});
  // Enter 觸發查詢
  ['txtItem','txtPO'].forEach(id => {{
    document.getElementById(id).addEventListener('keydown', e => {{
      if (e.key === 'Enter') doSearch();
    }});
  }});
  doSearch();
}}

function toInputDate(s) {{
  // YYYY/MM/DD → YYYY-MM-DD
  if (!s) return '';
  return s.replace(/\//g, '-');
}}

// ── 篩選 ───────────────────────────────────────────────────
function doSearch() {{
  const vendor   = document.getElementById('selVendor').value;
  const dateFrom = document.getElementById('dateFrom').value;
  const dateTo   = document.getElementById('dateTo').value;
  const txtItem  = document.getElementById('txtItem').value.trim().toLowerCase();
  const txtPO    = document.getElementById('txtPO').value.trim().toLowerCase();

  filteredData = RAW_DATA.filter(r => {{
    if (vendor && r['廠商簡稱'] !== vendor) return false;
    const d = toInputDate(r['驗收日期']);
    if (dateFrom && d < dateFrom) return false;
    if (dateTo   && d > dateTo)   return false;
    if (txtItem) {{
      const hay = (r['品號'] + ' ' + r['品名']).toLowerCase();
      if (!hay.includes(txtItem)) return false;
    }}
    if (txtPO) {{
      if (!(r['採購單號'] || '').toLowerCase().includes(txtPO)) return false;
    }}
    return true;
  }});

  applySort();
  currentPage = 1;
  renderTable();
  updateSortHeaders();
  updateSummary();
}}

function doClear() {{
  document.getElementById('selVendor').value = '';
  document.getElementById('txtItem').value   = '';
  document.getElementById('txtPO').value     = '';
  const dates = RAW_DATA.map(r => r['驗收日期']).filter(d => d).sort();
  if (dates.length) {{
    document.getElementById('dateFrom').value = toInputDate(dates[0]);
    document.getElementById('dateTo').value   = toInputDate(dates[dates.length - 1]);
  }}
  doSearch();
}}

// ── 排序 ───────────────────────────────────────────────────
function applySort() {{
  filteredData.sort((a, b) => {{
    let va = a[sortCol] ?? '';
    let vb = b[sortCol] ?? '';
    const na = parseFloat(va), nb = parseFloat(vb);
    if (!isNaN(na) && !isNaN(nb)) {{
      return sortAsc ? na - nb : nb - na;
    }}
    const sa = String(va), sb = String(vb);
    return sortAsc ? sa.localeCompare(sb, 'zh-TW') : sb.localeCompare(sa, 'zh-TW');
  }});
}}

function updateSortHeaders() {{
  document.querySelectorAll('thead th[data-col]').forEach(th => {{
    th.classList.remove('sort-asc', 'sort-desc');
    if (th.dataset.col === sortCol) {{
      th.classList.add(sortAsc ? 'sort-asc' : 'sort-desc');
    }}
  }});
}}

// ── 渲染表格 ───────────────────────────────────────────────
const NUM_COLS = new Set(['驗收數量','計價數量','單價','本幣未稅金額','本幣稅額','本幣金額合計']);
const DISP_COLS = ['驗收日期','廠商簡稱','品號','品名','規格','驗收數量','單位','單價','本幣未稅金額','本幣稅額','本幣金額合計','採購單號','發票號碼','備註'];

function fmtNum(v) {{
  if (v === '' || v === null || v === undefined) return '';
  const n = parseFloat(v);
  if (isNaN(n)) return v;
  if (Number.isInteger(n)) return n.toLocaleString('zh-TW');
  return n.toLocaleString('zh-TW', {{minimumFractionDigits: 0, maximumFractionDigits: 2}});
}}

function renderTable() {{
  const tbody = document.getElementById('tbody');
  const total = filteredData.length;
  const pages = Math.max(1, Math.ceil(total / PAGE_SIZE));
  if (currentPage > pages) currentPage = pages;

  const start = (currentPage - 1) * PAGE_SIZE;
  const slice = filteredData.slice(start, start + PAGE_SIZE);

  if (slice.length === 0) {{
    tbody.innerHTML = '<tr><td colspan="14" class="no-data">查無符合資料</td></tr>';
    renderPagination(0, 1);
    return;
  }}

  const rows = slice.map(r => {{
    const cells = DISP_COLS.map(col => {{
      const v = r[col] ?? '';
      if (col === '廠商簡稱') {{
        const color = getVendorColor(v);
        return `<td><span class="badge" style="background:${{color}}">${{v}}</span></td>`;
      }}
      if (NUM_COLS.has(col)) {{
        return `<td class="num">${{fmtNum(v)}}</td>`;
      }}
      return `<td>${{escHtml(String(v))}}</td>`;
    }});
    return `<tr>${{cells.join('')}}</tr>`;
  }});

  tbody.innerHTML = rows.join('');
  renderPagination(total, pages);
}}

function escHtml(s) {{
  return s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}}

// ── 分頁 ───────────────────────────────────────────────────
function renderPagination(total, pages) {{
  const pag = document.getElementById('pagination');
  if (total === 0) {{ pag.innerHTML = ''; return; }}

  const cp = currentPage;
  let html = '';

  html += `<button class="page-btn" onclick="goPage(${{cp-1}})" ${{cp===1?'disabled':''}}>&#8249;</button>`;

  const WINDOW = 2;
  let lo = Math.max(1, cp - WINDOW);
  let hi = Math.min(pages, cp + WINDOW);
  if (lo > 1) {{
    html += `<button class="page-btn" onclick="goPage(1)">1</button>`;
    if (lo > 2) html += `<span class="page-info">…</span>`;
  }}
  for (let i = lo; i <= hi; i++) {{
    html += `<button class="page-btn ${{i===cp?'active':''}}" onclick="goPage(${{i}})">${{i}}</button>`;
  }}
  if (hi < pages) {{
    if (hi < pages - 1) html += `<span class="page-info">…</span>`;
    html += `<button class="page-btn" onclick="goPage(${{pages}})">${{pages}}</button>`;
  }}

  html += `<button class="page-btn" onclick="goPage(${{cp+1}})" ${{cp===pages?'disabled':''}}>&#8250;</button>`;
  html += `<span class="page-info">第 ${{cp}} / ${{pages}} 頁，共 ${{total}} 筆</span>`;

  pag.innerHTML = html;
}}

function goPage(p) {{
  const pages = Math.max(1, Math.ceil(filteredData.length / PAGE_SIZE));
  if (p < 1 || p > pages) return;
  currentPage = p;
  renderTable();
  document.querySelector('.table-scroll').scrollTo({{top: 0, behavior: 'smooth'}});
}}

// ── 統計摘要 ───────────────────────────────────────────────
function updateSummary() {{
  let untax = 0, tax = 0, total = 0;
  filteredData.forEach(r => {{
    untax += parseFloat(r['本幣未稅金額']) || 0;
    tax   += parseFloat(r['本幣稅額'])    || 0;
    total += parseFloat(r['本幣金額合計']) || 0;
  }});
  document.getElementById('sumCount').textContent  = filteredData.length.toLocaleString('zh-TW');
  document.getElementById('sumUntax').textContent  = Math.round(untax).toLocaleString('zh-TW');
  document.getElementById('sumTax').textContent    = Math.round(tax).toLocaleString('zh-TW');
  document.getElementById('sumTotal').textContent  = Math.round(total).toLocaleString('zh-TW');
}}

// ── 啟動 ───────────────────────────────────────────────────
init();
</script>
</body>
</html>
"""

output_path = r'C:\Users\admin\OneDrive\桌面\進貨明細查詢.html'
with open(output_path, 'w', encoding='utf-8') as f:
    f.write(html)

print(f'HTML 已產生：{output_path}')
print(f'資料筆數：{len(records)}')
print(f'日期範圍：{date_min} ～ {date_max}')
