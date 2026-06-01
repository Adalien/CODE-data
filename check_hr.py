# -*- coding: utf-8 -*-
import openpyxl, sys, io, re
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

wb = openpyxl.load_workbook(r'C:\Users\admin\OneDrive\桌面\CODE資料\inv_tmp.xlsx', data_only=True)
ws = wb.active
rows = list(ws.iter_rows(values_only=True))

outer = []
skip_vals = {'熔噴','外層','親膚','耳線','需叫貨',''}
for row in rows:
    for ci, val in enumerate(row):
        if ci == 5: continue
        if not isinstance(val, str): continue
        v = val.strip()
        if v in skip_vals: continue
        m = re.match(r'^(\d+)\*(\d+)([^*]*)\*(.+)$', v)
        if not m: continue
        gsm   = int(m.group(1))
        width = int(m.group(2))
        kw    = m.group(3).strip()
        qty_s = m.group(4)
        if gsm < 28: continue

        wh   = sum(float(x) for x in re.findall(r'(\d+(?:\.\d+)?)[KkMmBb]', qty_s))
        s2   = re.sub(r'\d+(?:\.\d+)?[KkMmBb]', '', qty_s)
        m2   = re.search(r'(\d+(?:\.\d+)?)', s2)
        main = float(m2.group(1)) if m2 else 0
        qty  = main + wh
        outer.append((gsm, width, kw, qty, v))

# 由少到多排序
outer.sort(key=lambda x: x[3])

# 找昊瑞系列關鍵字
HR_KW = ['莫C','莫A','莫B','奶油','馬A','見A','見C','茶A','秋A','春A']

print('='*60)
print('【 外層布庫存 — 由少到多 】')
print('='*60)
print(f'  {"規格":<22} {"庫存":>5}  備注')
print('-'*60)

for gsm, width, kw, qty, raw in outer:
    is_hr = any(k in kw for k in HR_KW)
    hr_tag = ' ★昊瑞' if is_hr else ''
    if qty == 0:
        flag = ' 🔴 庫存0'
    elif qty <= 3:
        flag = ' ⚠️ 偏少'
    elif qty <= 8:
        flag = ' 注意'
    else:
        flag = ''
    print(f'  {gsm}g×{width}mm {kw:<12} {qty:>5.0f}捲{hr_tag}{flag}')
