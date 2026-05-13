# -*- coding: utf-8 -*-
import pandas as pd, warnings, sys, io
warnings.filterwarnings('ignore')
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

path = r"C:\Users\admin\OneDrive\桌面\四品號彙整_每週領出_20260301_20260508.xlsx"

# 讀總覽
df = pd.read_excel(path, sheet_name='總覽', header=2)
df.columns = ['週次','區間','BMB白175','BMB灰175','BES白175','BES白200','週合計']
df = df[df['週次'].astype(str).str.contains('W')].copy()
df['月'] = df['區間'].astype(str).str.extract(r'(\d{2})/\d{2}~')[0].astype(int)
for col in ['BES白175','BES白200']:
    df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

# 月合計（只取完整月：3月、4月）
monthly = df.groupby('月')[['BES白175','BES白200']].sum()
print("=== 親膚布每月領出量（捲）===")
print(f"  {'月份':<6} {'20×175':>8} {'20×200':>8}")
print(f"  {'-'*6} {'-'*8} {'-'*8}")
for m, row in monthly.iterrows():
    note = "（完整月）" if m in [3,4] else "（至5/8，不完整）"
    print(f"  {m}月     {row['BES白175']:>8.0f} {row['BES白200']:>8.0f}  {note}")

# 用完整月（3月、4月）計算月均
complete = monthly.loc[monthly.index.isin([3,4])]
avg175 = complete['BES白175'].mean()
avg200 = complete['BES白200'].mean()
avg260 = 42 / 2  # 使用者提供：兩個月42捲

print(f"\n=== 親膚布月平均領出量（以完整月3-4月為基準）===")
print(f"  20×175mm  BES-FM-W01-201753000  月均 {avg175:.1f} 捲")
print(f"  20×200mm  BES-FM-W01-202003000  月均 {avg200:.1f} 捲")
print(f"  20×260mm  BES-FM-W01-202603000  月均 {avg260:.1f} 捲  （實際2個月領出42捲）")
print(f"\n  月均合計: {avg175+avg200+avg260:.1f} 捲")

# 2個月預估需求 vs 現有庫存
BES_INV = {
    '20×175': 54.0,  # FM+JG合計
    '20×200': 90.0,
    '20×260': 40.0,
}
avgs = {'20×175': avg175, '20×200': avg200, '20×260': avg260}

print(f"\n=== 5-6月需求預估 vs 現有庫存 ===")
print(f"  {'尺寸':<10} {'月均':>6} {'2月需求':>8} {'現貨':>6} {'缺口':>8}")
print(f"  {'-'*10} {'-'*6} {'-'*8} {'-'*6} {'-'*8}")
t_need=0; t_gap=0
for size, a in avgs.items():
    need  = round(a*2, 1)
    stock = BES_INV[size]
    gap   = max(0, need-stock)
    t_need+=need; t_gap+=gap
    flag = "❌" if gap>0 else "✓ "
    print(f"  {flag} {size:<10} {a:>6.1f} {need:>8.1f} {stock:>6.1f} {gap:>8.1f}")
print(f"\n  {'合計':<12} {(avg175+avg200+avg260):>6.1f} {t_need:>8.1f} {sum(BES_INV.values()):>6.1f} {t_gap:>8.1f}")
print(f"\n  5-6月預估總需求: {t_need:.1f} 捲")
print(f"  需補購:          {t_gap:.1f} 捲")
