import sys, os, pandas as pd, io
sys.stdout.reconfigure(encoding='utf-8')

tmp = r'C:\Users\admin\AppData\Local\Temp\shortage_tmp.xlsx'
summ = pd.read_excel(tmp, sheet_name='材料需求彙總')
orders = pd.read_excel(tmp, sheet_name='訂單總覽')

# BOM
os.chdir(r'C:\Users\admin\OneDrive\桌面\Claude紀錄')
bom_files = [f for f in os.listdir('.') if f.startswith('BOM') and f.endswith('.xlsx')]
with open(sorted(bom_files)[-1], 'rb') as f:
    bom = pd.read_excel(io.BytesIO(f.read()))
c_mat = bom.columns[0]
c_prod = bom.columns[1]

# 目標BPP
targets = {
    '30×200 黑': ['BPP-HD-BK03302001500', 'BPP-SS-BK01302001500'],
    '40×260 黑': ['BPP-SS-BK01402601500'],
    '40×250 黑(TN95)': ['BPP-SS-BK01402501500'],
}

results = []

for spec, codes in targets.items():
    # 找對應成品品號
    prod_set = set()
    for code in codes:
        prods = bom[bom[c_mat].astype(str) == code][c_prod].dropna().tolist()
        prod_set.update(prods)

    # 需求 & 庫存（從材料需求彙總，多個code加總）
    rows_s = summ[summ['材料品號'].isin(codes)]
    total_demand_m = rows_s['總需求(m)'].sum()
    total_demand_roll = rows_s['總需求量'].sum()
    # 庫存：取唯一捲數（若兩個code共用庫存只取一個；但這裡直接加總）
    # 因為系統把同一批庫存分配給兩個code，所以取最大那個
    stock = rows_s['庫存量'].max()
    shortage = max(0, total_demand_roll - stock)

    # 找對應訂單
    for prod in sorted(prod_set):
        # 訂單總覽品號欄位做前綴匹配
        mask = orders['品號'].astype(str).apply(lambda x: x == prod or x.startswith(prod[:20]))
        matched = orders[mask]
        if matched.empty:
            # 嘗試去掉最後-XX後綴
            base = '-'.join(prod.split('-')[:4])
            mask2 = orders['品號'].astype(str).str.startswith(base)
            matched = orders[mask2]

        for _, r in matched.iterrows():
            results.append({
                '規格': spec,
                '品名': r['品名'],
                '品號': r['品號'],
                '訂單批號': r['訂單批號'],
                '生產量(盒)': int(r['生產量(盒)']) if pd.notna(r['生產量(盒)']) else 0,
                '需求(m)': round(total_demand_m),
                '需求(捲)': round(total_demand_roll, 2),
                '庫存(捲)': stock,
                '缺料(捲)': round(shortage, 2),
            })

df_out = pd.DataFrame(results).drop_duplicates(subset=['品名','訂單批號'])
print(df_out[['規格','品名','訂單批號','生產量(盒)','需求(捲)','庫存(捲)','缺料(捲)']].to_string(index=False))
