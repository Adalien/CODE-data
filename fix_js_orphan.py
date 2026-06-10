# -*- coding: utf-8 -*-
"""
清除 HTML 裡孤立的 JS 殘碼，修復 init() 無法執行的問題
"""
import sys, re, os, shutil, base64
sys.stdout.reconfigure(encoding='utf-8')

HTML_FILES = [
    r'C:\Users\admin\OneDrive\桌面\CODE資料\進貨明細查詢.html',
    r'C:\Users\admin\OneDrive\桌面\進貨明細查詢.html',
    r'C:\Users\admin\OneDrive\桌面\CODE資料\達特世生技_進貨明細查詢_2026.html',
    r'C:\Users\admin\OneDrive\桌面\達特世生技_進貨明細查詢_2026.html',
]
IMG_DIR = r'C:\Users\admin\OneDrive\桌面\CODE資料'
IMG_FILES = [
    ('1-150622200A8-50.jpg',                 'image/jpeg'),
    ('6a540aa3b8d6c1f69ef324319831a652.gif', 'image/gif'),
    ('CakeEmotion_0087830.gif',              'image/gif'),
    ('好想兔-賤.gif',                          'image/gif'),
]

SLIDESHOW_JS = (
    "\n// ── 賤兔輪播 ─────────────────────────────────────────────\n"
    "document.addEventListener('DOMContentLoaded',function(){\n"
    "  var cur=0;\n"
    "  setInterval(function(){\n"
    "    document.querySelectorAll('[id=\"filterMascot\"],[id^=\"mascot-tab-\"]').forEach(function(box){\n"
    "      box.querySelectorAll('img').forEach(function(img,i){img.style.opacity=(i===cur)?'1':'0';});\n"
    "    });\n"
    "    cur=(cur+1)%4;\n"
    "  },3000);\n"
    "});"
)

def fix_html(html):
    # 1. 找 init(); 的位置
    init_pos = html.find('init();')
    if init_pos < 0:
        print('  [WARN] init() not found')
        return html

    # 2. 找 init(); 之後，主 script 塊的 </script>
    #    從 init_pos 往後找第一個 </script>
    first_close = html.find('</script>', init_pos)
    if first_close < 0:
        print('  [WARN] </script> after init() not found')
        return html

    # 3. 主 script 塊：init(); 到 </script> 之間的孤立殘碼全部清除
    #    只保留 init(); 本身，加上乾淨的 SLIDESHOW_JS
    html = (html[:init_pos + len('init();')]
            + '\n' + SLIDESHOW_JS + '\n'
            + html[first_close:])

    # 4. 移除孤立的第二個 <script> 塊（殘碼 + 舊的 DOMContentLoaded）
    #    找到緊接在 </script> 後面（中間只有空白和HTML）的下一個 <script>...</script>
    #    這個塊只有殘碼，需要整個刪掉
    pattern = r'(</script>[\s\S]*?)<script>([\s\S]*?)</script>(\s*</body>)'
    def clean_extra_script(m):
        before = m.group(1)   # </script> + HTML between
        inner  = m.group(2)   # content of the extra script
        after  = m.group(3)   # </body>
        # 如果 inner 只有輪播/殘碼，刪掉這個 <script> 塊
        inner_stripped = inner.strip()
        if ('DOMContentLoaded' in inner_stripped or
            'cur = (cur+1)' in inner_stripped or
            'cur=(cur+1)' in inner_stripped or
            inner_stripped == ''):
            print('  [OK] 移除孤立的第二個 <script> 塊')
            return before + after
        return m.group(0)  # 保留

    html = re.sub(pattern, clean_extra_script, html, flags=re.DOTALL)

    return html


def verify(html, label=''):
    init_pos = html.find('init();')
    dcl = len(re.findall(r'DOMContentLoaded', html))
    # Count </script> after init
    scripts_after = html[init_pos:].count('</script>')
    orphan = bool(re.search(r'init\(\);\s*[\n\r]+\s*\}\)', html))
    print(f'  {label}: init_pos={init_pos}, DOMContentLoaded={dcl}, '
          f'</script> after init={scripts_after}, orphan_after_init={orphan}')


for path in HTML_FILES:
    if not os.path.exists(path):
        print(f'SKIP: {path}')
        continue
    print(f'\nProcessing: {os.path.basename(path)}')
    with open(path, 'r', encoding='utf-8') as f:
        html = f.read()
    verify(html, 'before')
    html_fixed = fix_html(html)
    verify(html_fixed, 'after ')
    with open(path, 'w', encoding='utf-8') as f:
        f.write(html_fixed)
    print(f'  Saved.')

print('\n[DONE]')
