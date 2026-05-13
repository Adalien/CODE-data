# -*- coding: utf-8 -*-
import pandas as pd, sys, io, warnings, os, glob
warnings.filterwarnings('ignore')
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

BASE = r"C:\Users\admin\OneDrive\桌面\CODE資料"
DESK = r"C:\Users\admin\OneDrive\桌面"

def latest2(*pats):
    files = []
    for pat in pats:
        files.extend(glob.glob(pat))
    return max(files, key=os.path.getmtime) if files else None

# 讀製令總表（歷史生產）
bom_path = latest2(os.path.join(BASE, "製令總表_合併*.xlsx"))
bom = pd.read_excel(bom_path)
bom.columns = [c.strip() for c in bom.columns]

mac = bom[bom['品名'].astype(str).str.contains('馬卡龍', na=False)].copy()
print(f"製令總表 馬卡龍四色相關製令：{len(mac)} 筆")
print()

mac['開單日期'] = pd.to_datetime(mac['開單日期'], errors='coerce')
mac['預計完工'] = pd.to_datetime(mac['預計完工'], errors='coerce')
mac = mac.sort_values('開單日期', ascending=False)

cols = ['製令單號','開單日期','預計完工','產品品號','品名','預計產量','已生產量','狀態碼']
print(mac[cols].to_string(index=False))

# 讀訂單總表（現有訂單）
order_path = latest2(os.path.join(DESK, "訂單生產總表*.xlsx"), os.path.join(BASE, "訂單生產總表*.xlsx"))
print(f"\n\n=== 訂單總表：{os.path.basename(order_path)} ===")
df = pd.read_excel(order_path, sheet_name='總表', header=0)
df.columns = ['訂單日期','序列','通路','品項','品號','品名','單位','機台','生產時間',
              '訂單量','庫存','差異','生產量','箱數','批號','製令單號','領料','入庫',
              '交期','出貨日','備註','機台1','生產時間1'][:len(df.columns)]

mac_order = df[df['品名'].astype(str).str.contains('馬卡龍', na=False)].copy()
print(f"訂單總表 馬卡龍相關訂單：{len(mac_order)} 筆")
mac_order['訂單日期'] = pd.to_datetime(mac_order['訂單日期'], errors='coerce')
mac_order = mac_order.sort_values('訂單日期', ascending=False)
print(mac_order[['訂單日期','序列','通路','品號','品名','訂單量','生產量','入庫','出貨日','備註']].to_string(index=False))
