import sys, os
sys.path.insert(0, os.path.abspath("."))
from app_tk13895 import parse_gl094_excel

open_bal, df = parse_gl094_excel("GL_094_TK13895 (010126-311026).xlsx")

mask_dnb = (
    df['dien_giai'].str.contains(r'điện mua nội bộ|điện mua nb|cp điện mua nội bộ|cp điện mua nb', case=False, na=False) |
    (df['dien_giai'].str.contains(r'khai thuế gtgt tháng 2/2026', case=False, na=False) & (df['co'] > 1e10)) |
    (df['dien_giai'].str.contains(r'hđ1335', case=False, na=False) & (df['co'].abs() > 1e10)) |
    (df['dien_giai'].str.contains(r'hạch toán cp điện mua nội bộ', case=False, na=False))
)
df_work = df[~mask_dnb].copy()

gl_groups = df_work[df_work['so_ct_gl'] != ''].groupby('so_ct_gl')
for gl, grp in gl_groups:
    if len(grp) >= 2 and (grp['no'] < 0).any() and abs(grp['no'].sum()) < 1.0:
        print(f"GL {gl}: {len(grp)} rows, sum No={grp['no'].sum()}, rows={grp['excel_row'].tolist()}")
