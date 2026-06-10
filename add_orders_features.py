# -*- coding: utf-8 -*-
"""
訂單總表功能升級：欄位篩選 + 品號點擊看 BOM
- 在訂單總表 thead 加入第二列篩選輸入框（16欄）
- 品號欄改為可點擊，顯示「此成品需要哪些原料」BOM 彈窗
- 注入相關 CSS 和 JS
"""
import sys, re, os
sys.stdout.reconfigure(encoding='utf-8')

HTML_FILES = [
    r'C:\Users\admin\OneDrive\桌面\CODE資料\進貨明細查詢.html',
    r'C:\Users\admin\OneDrive\桌面\進貨明細查詢.html',
    r'C:\Users\admin\OneDrive\桌面\CODE資料\達特世生技_進貨明細查詢_2026.html',
    r'C:\Users\admin\OneDrive\桌面\達特世生技_進貨明細查詢_2026.html',
]

# ── 訂單總表篩選 + BOM JS ──────────────────────────────────────
ORD_JS = r"""
// ── 訂單總表欄位篩選 ──────────────────────────────────────────
var ordColFilters = {};
function ordFilterInput(inp) {
  ordColFilters[inp.dataset.colidx] = inp.value.trim().toLowerCase();
  applyOrdFilter();
}
function applyOrdFilter() {
  document.querySelectorAll('#tab-orders tbody tr').forEach(function(row) {
    if (row.classList.contains('section-header')) { row.style.display = ''; return; }
    var show = true;
    Object.keys(ordColFilters).forEach(function(idx) {
      var kw = ordColFilters[idx];
      if (!kw) return;
      var td = row.cells[parseInt(idx)];
      if (td && !td.textContent.toLowerCase().includes(kw)) show = false;
    });
    row.style.display = show ? '' : 'none';
  });
}
// ── 訂單品號 BOM 彈窗（逆向查詢：成品 → 所需原料）─────────────
var _PBOM = null;
function _buildPBOM() {
  var nm = {};
  RAW_DATA.forEach(function(r) { if (r['品號'] && r['品名']) nm[r['品號']] = r['品名']; });
  var m = {};
  Object.keys(BOM_DATA).forEach(function(mc) {
    BOM_DATA[mc].forEach(function(p) {
      var pno = p['主件品號']; if (!m[pno]) m[pno] = [];
      m[pno].push({ mc: mc, mn: nm[mc] || '', qty: p['組成用量'], loss: p['損耗率'] });
    });
  });
  return m;
}
function showProductBOM(pno) {
  if (!_PBOM) _PBOM = _buildPBOM();
  var items = _PBOM[pno] || [];
  document.getElementById('bomTitle').textContent = '品號：' + pno;
  if (!items.length) {
    document.getElementById('bomSub').textContent = '此成品在 BOM 表中無對應原料';
    document.getElementById('bomContent').innerHTML =
      '<div class="bom-empty">查無 BOM 資料<br><small>可能為特殊訂製品或品號未建 BOM</small></div>';
  } else {
    document.getElementById('bomSub').textContent = '此成品共需 ' + items.length + ' 種原料：';
    document.getElementById('bomContent').innerHTML =
      '<table><thead><tr>' +
      '<th>原料品號</th><th>原料品名</th>' +
      '<th style="text-align:right">組成用量</th><th style="text-align:right">損耗率</th>' +
      '</tr></thead><tbody>' +
      items.map(function(r) {
        return '<tr><td>' + escHtml(r.mc) + '</td><td style="font-size:12px">' + escHtml(r.mn) +
               '</td><td class="num">' + (r.qty || 0) + '</td><td class="num">' + (r.loss || 0) + '%</td></tr>';
      }).join('') + '</tbody></table>';
  }
  document.getElementById('bomOverlay').classList.add('show');
}
"""

# ── 訂單品號 CSS ──────────────────────────────────────────────
ORD_CSS = """
  /* ── 訂單總表 品號可點擊 ── */
  td.ord-pno { cursor: pointer; color: #2563eb; text-decoration: underline; }
  td.ord-pno:hover { color: #1d4ed8; background: #eff6ff; }
"""

# 訂單總表 16 欄對應 index
ORD_COLS_N = 16


def patch_html(html, filename):
    changed = []

    # ── 1. 加篩選列到訂單總表 thead ──────────────────────────
    if 'ord-col-filter' not in html:
        filter_ths = ''.join(
            f'<th><input class="col-filter ord-col-filter" data-colidx="{i}"'
            f' placeholder="篩選" type="text" oninput="ordFilterInput(this)"></th>'
            for i in range(ORD_COLS_N)
        )
        filter_row = f'<tr class="col-filter-row">{filter_ths}</tr>'

        # 找訂單總表的 </tr></thead>（第一個出現在 tab-orders 之後的）
        orders_start = html.find('<div id="tab-orders"')
        if orders_start < 0:
            print(f'  [WARN] tab-orders not found')
        else:
            section = html[orders_start:]
            thead_end = section.find('</tr></thead>')
            if thead_end >= 0:
                insert_at = orders_start + thead_end + len('</tr>')
                html = html[:insert_at] + filter_row + html[insert_at:]
                changed.append('訂單篩選列')
            else:
                print(f'  [WARN] orders </tr></thead> not found')
    else:
        print(f'  [SKIP] ord-col-filter already exists')

    # ── 2. 讓訂單品號欄可點擊 ────────────────────────────────
    # 找訂單總表 tbody，把品號 <td> 改為 ord-pno
    if 'ord-pno' not in html:
        orders_start = html.find('<div id="tab-orders"')
        orders_end   = html.find('<!--/tab-orders-->', orders_start)
        if orders_start < 0 or orders_end < 0:
            print(f'  [WARN] tab-orders section not found')
        else:
            section = html[orders_start:orders_end]
            tbody_start = section.find('<tbody>')
            tbody_end   = section.find('</tbody>', tbody_start)
            if tbody_start < 0:
                print(f'  [WARN] tbody not found in tab-orders')
            else:
                tbody = section[tbody_start:tbody_end + 8]
                # 找每個 <tr> 的第5個 <td>（index 4, 品號欄），加 class + onclick
                # 策略：把 </td><td style="font-size:12px">{pno}</td><td style="font-size:12px"> 轉換
                # 用更直接的方法：逐列找第5個td

                def patch_tr_pno(m):
                    tr = m.group(0)
                    # 找所有 td 的起始位置
                    tds = list(re.finditer(r'<td[^>]*>', tr))
                    if len(tds) < 5:
                        return tr
                    td4 = tds[4]  # index 4 = 品號
                    td4_start = td4.start()
                    # 找這個 td 的結束 </td>
                    td4_end = tr.find('</td>', td4_start)
                    if td4_end < 0:
                        return tr
                    # 取出品號文字
                    pno_text = tr[td4.end():td4_end]
                    # 重建這個 td
                    new_td = f'<td class="ord-pno" style="font-size:12px" onclick="showProductBOM(\'{pno_text}\')">{pno_text}</td>'
                    return tr[:td4_start] + new_td + tr[td4_end + 5:]

                new_tbody = re.sub(
                    r'<tr[^>]*>.*?</tr>',
                    patch_tr_pno,
                    tbody,
                    flags=re.DOTALL
                )
                new_section = section[:tbody_start] + new_tbody[:-0] + section[tbody_end + 8:]
                # Rebuild: 這樣會把new_tbody整個替換
                new_section = section[:tbody_start] + new_tbody + section[tbody_end + 8:]
                html = html[:orders_start] + new_section + html[orders_end:]
                changed.append('品號 ord-pno onclick')
    else:
        print(f'  [SKIP] ord-pno already exists')

    # ── 3. 注入 ord-pno CSS ──────────────────────────────────
    if 'td.ord-pno' not in html:
        html = html.replace('</style>', ORD_CSS + '\n</style>', 1)
        changed.append('CSS')
    else:
        print(f'  [SKIP] CSS already exists')

    # ── 4. 注入 JS（在 function init() 前，只注入一次）────────
    if 'ordColFilters' not in html:
        init_fn = html.find('function init()')
        if init_fn >= 0:
            html = html[:init_fn] + ORD_JS + '\n' + html[init_fn:]
            changed.append('JS')
        else:
            print(f'  [WARN] function init() not found')
    else:
        print(f'  [SKIP] JS already exists')

    if changed:
        print(f'  [OK] 已加入: {", ".join(changed)}')
    return html


for path in HTML_FILES:
    if not os.path.exists(path):
        print(f'SKIP: {path}')
        continue
    print(f'\nProcessing: {os.path.basename(path)}')
    with open(path, 'r', encoding='utf-8') as f:
        html = f.read()
    html_new = patch_html(html, os.path.basename(path))
    with open(path, 'w', encoding='utf-8') as f:
        f.write(html_new)
    print(f'  Saved.')

print('\n[DONE]')
