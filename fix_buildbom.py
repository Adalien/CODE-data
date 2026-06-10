# -*- coding: utf-8 -*-
"""
修復 _buildPBOM 函式被截斷留下殘碼的問題
用括號計數法安全替換整個函式
"""
import sys, re, os
sys.stdout.reconfigure(encoding='utf-8')

HTML_FILES = [
    r'C:\Users\admin\OneDrive\桌面\CODE資料\進貨明細查詢.html',
    r'C:\Users\admin\OneDrive\桌面\進貨明細查詢.html',
    r'C:\Users\admin\OneDrive\桌面\CODE資料\達特世生技_進貨明細查詢_2026.html',
    r'C:\Users\admin\OneDrive\桌面\達特世生技_進貨明細查詢_2026.html',
]

NEW_BUILD_PBOM = """function _buildPBOM() {
  var nm = {};
  if (typeof MAT_NAME_DATA !== 'undefined') {
    Object.keys(MAT_NAME_DATA).forEach(function(k) { nm[k] = MAT_NAME_DATA[k]; });
  }
  RAW_DATA.forEach(function(r) { if (r['品號'] && r['品名']) nm[r['品號']] = r['品名']; });
  var m = {};
  Object.keys(BOM_DATA).forEach(function(mc) {
    BOM_DATA[mc].forEach(function(p) {
      var pno = p['主件品號']; if (!m[pno]) m[pno] = [];
      m[pno].push({ mc: mc, mn: nm[mc] || mc, qty: p['組成用量'], loss: p['損耗率'] });
    });
  });
  return m;
}"""


def replace_fn_safe(html, fn_name, new_fn):
    """用括號計數找函式完整範圍，安全替換"""
    marker = f'function {fn_name}('
    start = html.find(marker)
    if start < 0:
        return html, False
    brace_start = html.find('{', start)
    if brace_start < 0:
        return html, False
    depth = 0
    pos = brace_start
    in_str = None
    while pos < len(html):
        c = html[pos]
        # 簡單字串跳過（不處理跳脫，但足夠用）
        if in_str:
            if c == in_str and (pos == 0 or html[pos-1] != '\\'):
                in_str = None
        elif c in ('"', "'", '`'):
            in_str = c
        elif c == '{':
            depth += 1
        elif c == '}':
            depth -= 1
            if depth == 0:
                end = pos + 1
                return html[:start] + new_fn + html[end:], True
        pos += 1
    return html, False


def find_fn_end(html, fn_name):
    """用括號計數找函式結束位置（closing } 後的 index）"""
    marker = f'function {fn_name}('
    start = html.find(marker)
    if start < 0:
        return -1, -1
    brace_start = html.find('{', start)
    if brace_start < 0:
        return -1, -1
    depth = 0
    pos = brace_start
    in_str = None
    while pos < len(html):
        c = html[pos]
        if in_str:
            if c == in_str and (pos == 0 or html[pos-1] != '\\'):
                in_str = None
        elif c in ('"', "'", '`'):
            in_str = c
        elif c == '{':
            depth += 1
        elif c == '}':
            depth -= 1
            if depth == 0:
                return start, pos + 1   # (fn_start, fn_end)
        pos += 1
    return start, -1


def fix_html(html):
    marker = 'function _buildPBOM('
    fn_start = html.find(marker)
    if fn_start < 0:
        print('  [WARN] _buildPBOM 找不到')
        return html

    # 找 _buildPBOM 的結束：先嘗試括號計數法
    _, fn_end = find_fn_end(html, '_buildPBOM')

    if fn_end < 0:
        # 括號計數失敗（函式已截斷/損毀）→ 直接找 \nfunction showProductBOM 當結束邊界
        print('  [INFO] _buildPBOM 已截斷，使用 showProductBOM 邊界修復')
        fn_end_by_next = html.find('\nfunction showProductBOM(', fn_start)
        if fn_end_by_next < 0:
            print('  [WARN] showProductBOM 也找不到，無法修復')
            return html
        fn_end = fn_end_by_next  # 取代到這裡（不含）
    else:
        # 括號計數成功 → 還需清除後面的殘碼
        next_fn = html.find('\nfunction showProductBOM(', fn_end)
        if next_fn < 0:
            print('  [WARN] showProductBOM 找不到')
            return html
        between = html[fn_end:next_fn]
        if between.strip(' \t\n;'):
            print(f'  [INFO] 清除殘碼: {repr(between[:80])}')
            # 更新 fn_end 到 showProductBOM 之前
            fn_end = next_fn  # 取代到這裡（不含）
        else:
            # 無殘碼，只要替換函式本身
            html_new = html[:fn_start] + NEW_BUILD_PBOM + html[fn_end:]
            print('  [OK] _buildPBOM 已替換（無殘碼）')
            _check(html_new)
            return html_new

    # 統一：從 fn_start 到 fn_end（showProductBOM 前）整個替換
    html_new = html[:fn_start] + NEW_BUILD_PBOM + '\n' + html[fn_end + 1:]  # +1 跳過 \n
    print('  [OK] _buildPBOM 已修復')
    _check(html_new)
    return html_new


def _check(html):
    if "nm[mc] || ''" in html:
        print("  [WARN] 仍有舊版殘碼 nm[mc] || ''！")
    else:
        print('  [OK] 無舊版殘碼')
    # 驗證函式是否完整
    _, end = find_fn_end(html, '_buildPBOM')
    if end > 0:
        print(f'  [OK] _buildPBOM 完整（結尾 pos={end}）')
    else:
        print('  [WARN] _buildPBOM 仍有問題')


for path in HTML_FILES:
    if not os.path.exists(path):
        print(f'SKIP: {path}')
        continue
    print(f'\nProcessing: {os.path.basename(path)}')
    with open(path, 'r', encoding='utf-8') as f:
        html = f.read()
    html_new = fix_html(html)
    with open(path, 'w', encoding='utf-8') as f:
        f.write(html_new)
    print(f'  Saved.')

print('\n[DONE]')
