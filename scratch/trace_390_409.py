import sys, os
sys.path.insert(0, os.path.abspath("."))
from app_tk13895 import parse_gl094_excel, reconcile_tk13895

open_bal, df_raw = parse_gl094_excel("GL_094_TK13895 (010126-311026).xlsx")

# Check Row 390 and 409 in df_raw
r390 = df_raw[df_raw['excel_row'] == 390].iloc[0]
r409 = df_raw[df_raw['excel_row'] == 409].iloc[0]

print("r390:", r390['excel_row'], r390['so_hd'], r390['no'], r390['co'])
print("r409:", r409['excel_row'], r409['so_hd'], r409['no'], r409['co'])

open_bal, df, df_ton, df_da_gt, df_2025, df_dnb, df_t1_netoff, df_mistyped = reconcile_tk13895(df_raw, open_bal)

print("In df_ton:", 390 in df_ton['excel_row'].values)
print("In df_2025:", 409 in df_2025['excel_row'].values)
print("In df_t1_netoff:", 390 in df_t1_netoff['excel_row'].values, 409 in df_t1_netoff['excel_row'].values)
print("In df_da_gt:", 390 in df_da_gt['excel_row'].values, 409 in df_da_gt['excel_row'].values)
