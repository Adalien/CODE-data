import os, base64, re, sys, shutil
sys.stdout.reconfigure(encoding='utf-8')

HTML_FILES = [
    r'C:\Users\admin\OneDrive\桌面\CODE資料\進貨明細查詢.html',
    r'C:\Users\admin\OneDrive\桌面\達特世生技_進貨明細查詢_2026.html',
    r'C:\Users\admin\OneDrive\桌面\CODE資料\達特世生技_進貗明細查詢_2026.html',
]
IMG_DIR = r'C:\Users\admin\OneDrive\桌面\CODE資料'
IMG_FILES = [
    ('1-150622200A8-50.jpg',                    'image/jpeg'),
    ('6a540aa3b8d6c1f69ef324319831a652.gif',    'image/gif'),
    ('CakeEmotion_0087830.gif',                 'image/gif'),
    ('好想兔-賤.gif',                             'image/gif'),
]

# Build base64 srcs
srcs = []
for fname, mime in IMG_FILES:
    path = os.path.join(IMG_DIR, fname)
    with open(path, 'rb') as f:
        data = base64.b64encode(f.read()).decode('ascii')
    srcs.append(f'data:{mime};base64,{data}')

IMG_STYLE = "position:absolute;inset:0;width:100%;height:100%;object-fit:contain;transition:opacity 0.3s;"

def make_img_tags():
    tags = []
    for i, src in enumerate(srcs):
        opacity = '1' if i == 0 else '0'
        tags.append(f'<img src="{src}" style="{IMG_STYLE}opacity:{opacity};">')
    return '\n  '.join(tags)

def make_other_tab_mascot(tab_id):
    inner = make_img_tags()
    return (
        f'<div style="float:right;margin:0 18px 0 0;padding-top:4px;">'
        f'<div id="mascot-{tab_id}" style="width:260px;height:110px;position:relative;border-radius:8px;overflow:hidden;">'
        f'\n  {inner}\n'
        f'</div></div>\n'
    )

def make_filter_mascot():
    inner = make_img_tags()
    return (
        f'<div id="filterMascot" style="flex:0 0 auto;width:260px;height:110px;position:relative;border-radius:8px;overflow:hidden;">'
        f'\n  {inner}\n'
        f'</div>'
    )

# JS slideshow (single block)
SLIDESHOW_JS = '''
// ── 賤兔輪播 ──────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', function(){
  var cur = 0;
  setInterval(function(){
    document.querySelectorAll('[id="filterMascot"],[id^="mascot-tab-"]').forEach(function(box){
      var imgs = box.querySelectorAll('img');
      imgs.forEach(function(img, i){ img.style.opacity = (i===cur)?'1':'0'; });
    });
    cur = (cur+1) % 4;
  }, 3000);
});
'''

OTHER_TABS = ['tab-orders', 'tab-shortage', 'tab-inventory', 'tab-material', 'tab-pick', 'tab-ready']

def fix_html(html):
    # 1. Remove ALL existing mascot divs from other tabs (mascot-tab-*)
    # Pattern: <div style="float:right..."><div id="mascot-tab-...">...</div></div>
    for tab_id in OTHER_TABS:
        mid = f'mascot-{tab_id}'
        # Find and remove the outer float:right wrapper
        pattern = r'<div style="float:right[^>]*>\s*<div id="' + re.escape(mid) + r'"[^>]*>.*?</div>\s*</div>\s*'
        html = re.sub(pattern, '', html, flags=re.DOTALL)

    # 2. Remove ALL existing filterMascot divs
    # Keep track of how many we find
    fm_pattern = r'<div[^>]+id="filterMascot"[^>]*>.*?</div>'
    fm_matches = list(re.finditer(fm_pattern, html, re.DOTALL))
    # We'll rebuild filterMascot fresh

    # Remove all filterMascot divs
    html = re.sub(r'<div[^>]+id="filterMascot"[^>]*>.*?</div>', '', html, flags=re.DOTALL)

    # 3. Add fresh filterMascot in filter-row (after 清除 button)
    # Find: </button>\n    </div>\n    <div class="btn-row"> or similar pattern with 清除
    # The 清除 button is: onclick="doClear()"
    clear_btn_end = html.find('onclick="doClear()"')
    if clear_btn_end >= 0:
        # Find the end of this button tag
        btn_end = html.find('</button>', clear_btn_end)
        if btn_end >= 0:
            insert_pos = btn_end + len('</button>')
            mascot = '\n    ' + make_filter_mascot()
            html = html[:insert_pos] + mascot + html[insert_pos:]
    else:
        print('[WARN] 清除 button not found')

    # 4. Add mascot to other tabs - insert right after the opening div tag
    for tab_id in OTHER_TABS:
        pattern = f'<div id="{tab_id}" class="tab-content">'
        pos = html.find(pattern)
        if pos >= 0:
            insert_pos = pos + len(pattern)
            mascot = '\n' + make_other_tab_mascot(tab_id)
            html = html[:insert_pos] + mascot + html[insert_pos:]
            print(f'[OK] Added mascot to {tab_id}')
        else:
            print(f'[WARN] {tab_id} not found')

    # 5. Remove ALL DOMContentLoaded slideshow blocks
    html = re.sub(
        r"//\s*──\s*賤兔輪播\s*──+\s*document\.addEventListener\('DOMContentLoaded'.*?\}\);",
        '', html, flags=re.DOTALL
    )
    # Also remove bare DOMContentLoaded mascot blocks (without the comment header)
    html = re.sub(
        r"document\.addEventListener\('DOMContentLoaded',\s*function\(\)\{\s*var cur = 0;\s*setInterval.*?\}\);\s*\}\);",
        '', html, flags=re.DOTALL
    )

    # 6. Add single fresh slideshow JS before </script> closing (find last </script>)
    # Insert before the last </script> in the main script block
    last_script_close = html.rfind('</script>')
    if last_script_close >= 0:
        html = html[:last_script_close] + SLIDESHOW_JS + '\n' + html[last_script_close:]
    else:
        print('[WARN] </script> not found')

    return html


for html_path in HTML_FILES:
    if not os.path.exists(html_path):
        print(f'SKIP (not found): {html_path}')
        continue
    print(f'\nProcessing: {html_path}')
    with open(html_path, 'r', encoding='utf-8') as f:
        html = f.read()
    html_fixed = fix_html(html)
    with open(html_path, 'w', encoding='utf-8') as f:
        f.write(html_fixed)

    # Verify
    fm_count = len(re.findall(r'id="filterMascot"', html_fixed))
    dcl_count = len(re.findall(r'DOMContentLoaded', html_fixed))
    tab_counts = {}
    for tid in OTHER_TABS:
        mid = f'mascot-{tid}'
        cnt = len(re.findall(re.escape(f'id="{mid}"'), html_fixed))
        tab_counts[tid] = cnt
    print(f'  filterMascot occurrences: {fm_count} (expect 1+ in JS strings)')
    print(f'  DOMContentLoaded: {dcl_count}')
    print(f'  Other tab mascots: {tab_counts}')

print('\n[DONE]')
