# -*- coding: utf-8 -*-
import pandas as pd, sys, io, warnings
warnings.filterwarnings('ignore')
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

BASE = r"C:\Users\admin\OneDrive\桌面\CODE資料"
import glob, os

master_path = max(glob.glob(os.path.join(BASE, "BOM*.xlsx")), key=os.path.getmtime)
master = pd.read_excel(master_path)
master.columns = [c.strip() for c in master.columns]

# 找所有馬A紫相關品號
bpp = master[master['元件品號'].astype(str).str.contains('BPP', na=False)].copy()
maa = bpp[bpp['品     名'].astype(str).str.contains('馬', na=False) |
          bpp['元件品號'].astype(str).str.contains('PR17', na=False)]

print("=== 馬A紫 相關品號 ===")
print(maa[['元件品號','品     名','規     格']].drop_duplicates('元件品號').to_string(index=False))

print("\n=== 馬A紫 用在哪些成品 ===")
used = maa[maa['主件品號'].notna() & (maa['主件品號'].astype(str) != 'nan')]
print(used[['元件品號','主件品號','品     名','組成用量','底數']].to_string(index=False))
