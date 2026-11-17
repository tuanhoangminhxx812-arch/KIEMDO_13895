import openpyxl
import pandas as pd
import numpy as np
from app_tk13895 import parse_gl094_excel, reconcile_tk13895

open_bal, df = parse_gl094_excel('GL_094_TK13895 (010126-311026).xlsx')
open_bal, df_raw, df_ton, df_da_gt, df_2025, df_dnb, df_t1_netoff, df_mistyped = reconcile_tk13895(df, open_bal)

print(f"Total rows in df_raw: {len(df_raw)}")
print(f"df_ton rows: {len(df_ton)}, sum: {df_ton['no'].sum():,}")
print(f"Open balance: {open_bal:,}")

# Let's inspect rem_2025:
# Dòng 11 có số dư đầu kỳ: 125,834,024,708
# Các khoản Có giải trừ cho 2025 trong sổ:
# - Các dòng Có trong df_2025:
sum_2025_work = df_2025['co'].sum()
print(f"df_2025 Co sum: {sum_2025_work:,}")

# - Các dòng Có trong df_dnb giải trừ cho 2025:
# In df_dnb, let's see which Co lines relate to 2025:
# In 2026, DNB No has:
# Row 21 (T1/26): 96,675,410,321 -> giải trừ ở Row 50: 132,700,116,607 (trong đó 96,675 tỷ là của T1/26, 36,024 tỷ là của 2025)
# Row 18 (T1/26 kê khai): 87,840,801,042 (của HĐ 6 ngày 06/01/26 - tiền điện T12/2025)
# Tổng Co DNB của 2025 = 87,840,801,042 + (132,700,116,607 - 96,675,410,321) = 123,865,507,328
# Và df_2025 sum = 1,267,428,080
# Tổng Co giải trừ 2025 = 123,865,507,328 + 1,267,428,080 = 125,132,935,408
# Số dư đầu kỳ 2025 còn tồn = 125,834,024,708 - 125,132,935,408 = 701,089,300.
rem_2025 = 701089300.0

# 4 Components:
c1_ton_2026 = df_ton['no'].sum()
c2_dnb_t10 = 99355802395.0
c3_lech_563 = 30000000.0
c4_rem_2025 = rem_2025

total_ton_thuc_te = c1_ton_2026 + c2_dnb_t10 + c3_lech_563 + c4_rem_2025
close_bal = 102081918814.0

print(f"C1 - 60 HD ton 2026:      {c1_ton_2026:18,f}")
print(f"C2 - DNB T10 (row 564):    {c2_dnb_t10:18,f}")
print(f"C3 - Lech 563 (row 99):     {c3_lech_563:18,f}")
print(f"C4 - Du dau ky 2025 con:   {c4_rem_2025:18,f}")
print(f"---------------------------------------------------")
print(f"TONG TIEN CON TON THUC TE: {total_ton_thuc_te:18,f}")
print(f"SO DU CUOI KY TREN SO:     {close_bal:18,f}")
print(f"CHENH LECH:                {total_ton_thuc_te - close_bal:18,f}")
