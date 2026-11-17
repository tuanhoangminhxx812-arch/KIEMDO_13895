import sys, os
sys.path.insert(0, os.path.abspath("."))
import pandas as pd
from app_tk13895 import parse_gl094_excel

open_bal, df_work = parse_gl094_excel("GL_094_TK13895 (010126-311026).xlsx")

# Run Tầng 1 with fix
mask_dnb = (
    df_work['dien_giai'].str.contains(r'điện mua nội bộ|điện mua nb|cp điện mua nội bộ|cp điện mua nb', case=False, na=False) |
    (df_work['dien_giai'].str.contains(r'khai thuế gtgt tháng 2/2026', case=False, na=False) & (df_work['co'] > 1e10)) |
    (df_work['dien_giai'].str.contains(r'hđ1335', case=False, na=False) & (df_work['co'].abs() > 1e10)) |
    (df_work['dien_giai'].str.contains(r'hạch toán cp điện mua nội bộ', case=False, na=False))
)
df_dnb = df_work[mask_dnb].copy()
df_w = df_work[~mask_dnb].copy()

drop_indices_t1 = set()

# 1a. Trực tiếp cùng số CT phụ AP (+No / -No)
ap_groups = df_w[df_w['so_ct_phu'] != ''].groupby('so_ct_phu')
for ap, grp in ap_groups:
    if len(grp) >= 2:
        pos = grp[grp['no'] > 0]
        neg = grp[grp['no'] < 0]
        for _, r_neg in neg.iterrows():
            amt = abs(r_neg['no'])
            match_pos = pos[(pos['no'] == amt) & (~pos.index.isin(drop_indices_t1))]
            if len(match_pos) > 0:
                drop_indices_t1.add(r_neg.name)
                drop_indices_t1.add(match_pos.index[0])

# 1b. Bút toán điều chỉnh cấp đối tượng (GL 1366 sum == 0) - MUST exclude already dropped rows!
gl_groups = df_w[(df_w['so_ct_gl'] != '') & (~df_w.index.isin(drop_indices_t1))].groupby('so_ct_gl')
for gl, grp in gl_groups:
    if len(grp) >= 2 and (grp['no'] < 0).any() and abs(grp['no'].sum()) < 1.0:
        drop_indices_t1.update(grp.index.tolist())

print("Rows in drop_indices_t1 containing 388, 389, 390, 409:")
for r in [388, 389, 390, 409]:
    idx = df_w[df_w['excel_row'] == r].index
    if len(idx) > 0:
        print(f"Row {r} in drop_indices_t1?: {idx[0] in drop_indices_t1}")
