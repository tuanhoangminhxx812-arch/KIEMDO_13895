import sys, os
sys.path.insert(0, os.path.abspath("."))
import pandas as pd
from app_tk13895 import parse_gl094_excel, reconcile_tk13895

open_bal, df_raw = parse_gl094_excel("GL_094_TK13895 (010126-311026).xlsx")
open_bal, df, df_ton, df_da_gt, df_2025, df_dnb, df_t1_netoff, df_mistyped = reconcile_tk13895(df_raw, open_bal)

print("=== df_dnb Có > 0 ===")
for _, r in df_dnb[df_dnb['co'] > 0].iterrows():
    print(f"Row {r['excel_row']} ({r['ngay']}): Có={r['co']:,.0f} | Diễn giải: {r['dien_giai']}")
