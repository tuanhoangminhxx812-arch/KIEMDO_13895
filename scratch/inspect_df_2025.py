import sys, os
sys.path.insert(0, os.path.abspath("."))
import pandas as pd
from app_tk13895 import parse_gl094_excel, reconcile_tk13895

open_bal, df_raw = parse_gl094_excel("GL_094_TK13895 (010126-311026).xlsx")
open_bal, df, df_ton, df_da_gt, df_2025, df_dnb, df_t1_netoff, df_mistyped = reconcile_tk13895(df_raw, open_bal)

print(f"Total rows in df_2025: {len(df_2025)}")
print(f"Total Có in df_2025: {df_2025['co'].sum():,.0f} đ")
for _, r in df_2025.iterrows():
    print(f"Row {r['excel_row']} ({r['ngay']}): Có={r['co']:,.0f} | HĐ={r['so_hd']} | {r['dien_giai']}")
