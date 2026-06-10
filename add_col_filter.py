# -*- coding: utf-8 -*-
"""
為進貨明細表格欄位標題加入篩選輸入框
"""
import sys, re, os
sys.stdout.reconfigure(encoding='utf-8')

HTML_FILES = [
    r'C:\Users\admin\OneDrive\桌面\CODE資料\進貨明細查詢.html',
    r'C:\Users\admin\OneDrive\桌面\進貨明細查詢.html',
    r'C:\Users\admin\OneDrive\桌面\CODE資料\達特世生技_進貨明細查詢_2026.html',
    r'C:\Users\admin\OneDrive\桌面\達特世生技_進貨明細查詢_2026.html',
]

# ── 欄位篩選列 HTML（緊接在原本的 <tr> 後面）────────────────
COLS = [
    ('驗收日期', False),
    ('廠商簡稱', False),
    ('品號',     False),
    ('品名',     False),
    ('規格',     False),
    ('驗收數量', True),
    ('單位',     False),
    ('單價',     True),
    ('本幣未稅金額', True),
    ('本幣稅額',    True),
    ('本幣金額合計', True),
    ('採購單號', False),
    ('發票號碼', False),
    ('備註',     False),
]

filter_ths = '\n'.join(
    f'          <th class="{"num" if is_num else ""}">'
    f'<input class="col-filter" data-col="{col}" placeholder="篩選" type="text">'
    f'</th>'
    for col, is_num in COLS
)
FILTER_ROW = f'''      </tr>
      <tr class="col-filter-row">
{filter_ths}
      </tr>'''

# ── 新增 CSS ────────────────────────────────────────────────
NEW_CSS = '''
  /* ── 欄位篩選輸入框 ── */
  .col-filter-row th {
    padding: 4px 6px;
    background: #dce8ff;
  }
  .col-filter {
    width: 100%;
    box-sizing: border-box;
    padding: 3px 6px;
    border: 1px solid #b0c4e8;
    border-radius: 4px;
    font-size: 12px;
    font-weight: normal;
    background: #fff;
    color: #2d3748;
  }
  .col-filter:focus {
    outline: none;
    border-color: #2563eb;
    box-shadow: 0 0 0 2px rgba(37,99,235,0.15);
  }
'''

# ── 新增 JS（colFilters + 修改 doSearch）─────────────────────
# 插在 var colFilters = {} 宣告 + event listeners
INIT_COLFILTER_JS = '''
  // ── 欄位篩選 ────────────────────────────────────────────
  document.querySelectorAll('.col-filter').forEach(function(inp){
    inp.addEventListener('input', function(){
      colFilters[inp.dataset.col] = inp.value.trim().toLowerCase();
      currentPage = 1;
      applyColFilter();
    });
  });
'''

# colFilters 宣告 — 插在 filteredData 宣告後
COL_FILTER_VAR = 'var colFilters = {};\n'

# applyColFilter 函式 — 替換 doSearch 之後不動原 doSearch
APPLY_COL_FILTER_FN = '''
function applyColFilter() {
  var base = RAW_DATA.slice();
  // 先套用頂部篩選條件（同 doSearch 邏輯）
  var vendor   = document.getElementById('selVendor').value;
  var dateFrom = document.getElementById('dateFrom').value;
  var dateTo   = document.getElementById('dateTo').value;
  var txtItem  = document.getElementById('txtItem').value.trim().toLowerCase();
  var txtPO    = document.getElementById('txtPO').value.trim().toLowerCase();
  base = base.filter(function(r){
    if (vendor && r['廠商簡稱'] !== vendor) return false;
    var d = toInputDate(r['驗收日期']);
    if (dateFrom && d < dateFrom) return false;
    if (dateTo   && d > dateTo)   return false;
    if (txtItem) {
      var hay = ((r['品號']||'') + ' ' + (r['品名']||'')).toLowerCase();
      if (!hay.includes(txtItem)) return false;
    }
    if (txtPO && !((r['採購單號']||'').toLowerCase().includes(txtPO))) return false;
    return true;
  });
  // 再套用欄位篩選
  base = base.filter(function(r){
    return Object.keys(colFilters).every(function(col){
      var kw = colFilters[col];
      if (!kw) return true;
      return String(r[col]||'').toLowerCase().includes(kw);
    });
  });
  filteredData = base;
  applySort();
  renderTable();
  updateSortHeaders();
  updateSummary();
}
'''

def patch_html(html):
    # 1. 加篩選列 HTML（只加一次）
    if 'col-filter-row' in html:
        print('  [SKIP] col-filter-row already exists')
    else:
        # 在第一個 </tr>\n      </thead> 前插入篩選列
        html = html.replace(
            '      </tr>\n      </thead>',
            FILTER_ROW + '\n      </thead>',
            1
        )
        print('  [OK] 欄位篩選列已加入 thead')

    # 2. 加 CSS
    if '.col-filter' not in html:
        html = html.replace('</style>', NEW_CSS + '\n</style>', 1)
        print('  [OK] CSS 已加入')

    # 3. 加 colFilters 宣告（在 var filteredData 後）
    if 'colFilters' not in html:
        html = html.replace(
            'var filteredData = [];',
            'var filteredData = [];\n' + COL_FILTER_VAR,
            1
        )
        print('  [OK] colFilters 宣告已加入')

    # 4. 加 applyColFilter 函式（在 doSearch 函式後）
    if 'applyColFilter' not in html:
        # 找 doSearch 函式的結尾
        m = re.search(r'(function doSearch\(\).*?\n\})', html, re.DOTALL)
        if m:
            html = html[:m.end()] + '\n' + APPLY_COL_FILTER_FN + html[m.end():]
            print('  [OK] applyColFilter 函式已加入')

    # 5. 在 init() 裡加欄位篩選 event listeners（在 doSearch(); 前）
    if "inp.addEventListener('input'" not in html:
        html = html.replace(
            '  doSearch();\n}',
            INIT_COLFILTER_JS + '  doSearch();\n}',
            1
        )
        print('  [OK] 欄位篩選 event listeners 已加入 init()')

    return html


for path in HTML_FILES:
    if not os.path.exists(path):
        print(f'SKIP: {path}')
        continue
    print(f'\nProcessing: {os.path.basename(path)}')
    with open(path, 'r', encoding='utf-8') as f:
        html = f.read()
    html_new = patch_html(html)
    with open(path, 'w', encoding='utf-8') as f:
        f.write(html_new)
    print(f'  Saved.')

print('\n[DONE]')
