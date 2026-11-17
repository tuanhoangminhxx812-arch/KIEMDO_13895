import sys, os
sys.path.insert(0, os.path.abspath("."))
import pandas as pd
from app_tk13895 import parse_gl094_excel, reconcile_tk13895

open_bal, df_raw = parse_gl094_excel("GL_094_TK13895 (010126-311026).xlsx")
open_bal, df, df_ton, df_da_gt, df_2025, df_dnb, df_t1_netoff, df_mistyped = reconcile_tk13895(df_raw, open_bal)

print(f"Số dư đầu kỳ: {open_bal:,.0f} đ")
sum_co_2025 = df_2025['co'].sum()
print(f"Tổng Có giải trừ HĐ 2025 (df_2025): {sum_co_2025:,.0f} đ ({len(df_2025)} dòng)")

# Check điện mua nội bộ T12/2025 giải trừ tại T1 và T2
dnb_t12_2025 = df_dnb[(df_dnb['co'] > 0) & (df_dnb['dien_giai'].str.contains(r'2025|12/25|tháng 12', case=False, na=False))]
print(f"Điện mua nội bộ T12/2025 giải trừ bên Có: {dnb_t12_2025['co'].sum():,.0f} đ")
for _, r in dnb_t12_2025.iterrows():
    print(f"  Row {r['excel_row']}: {r['dien_giai']} -> Có: {r['co']:,.0f}")

total_co_for_2025 = sum_co_2025 + dnb_t12_2025['co'].sum()
print(f"Tổng cộng Có giải trừ cho số dư 2025: {total_co_for_2025:,.0f} đ")
rem_2025 = open_bal - total_co_for_2025
print(f"Số dư nợ 2025 còn lại chưa kê khai giải trừ: {rem_2025:,.0f} đ")
