import io, sys, os
sys.path.insert(0, os.path.abspath("."))
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from app_tk13895 import parse_gl094_excel, reconcile_tk13895

open_bal, df_raw = parse_gl094_excel("GL_094_TK13895 (010126-311026).xlsx")
open_bal, df, df_ton, df_da_gt, df_2025, df_dnb, df_t1_netoff, df_mistyped = reconcile_tk13895(df_raw, open_bal)

# Build 2025 pending dataframe
df_ton_2025 = pd.DataFrame([
    {
        'STT': 1,
        'Nhom_Ton': 'Tồn năm 2025',
        'excel_row': 205,
        'ngay': '31/05/2026',
        'so_ct_phu': '25008',
        'so_ct_gl': '1366',
        'so_hd': '3704 & 3705',
        'dien_giai': '25008 Phải trả VAT HD3704 và HĐ3705 ngày 28.07.25 cty Cadivi_Điều chỉnh cấp đối tượng TK13895_9000',
        'no': 76238400.0,
        'nguoi_lap': 'HIENNTK'
    },
    {
        'STT': 2,
        'Nhom_Ton': 'Tồn năm 2025',
        'excel_row': 206,
        'ngay': '31/05/2026',
        'so_ct_phu': '25017',
        'so_ct_gl': '1366',
        'so_hd': '102 & 691',
        'dien_giai': '25017 Phải trả VAT HD0102 ngày 05.08.25 cty E.E và 25017 Phải trả VAT HD0691 ngày 11.08.25 cty Hanaka_Điều chỉnh cấp đối tượng TK13895_9000',
        'no': 299911123.0,
        'nguoi_lap': 'HIENNTK'
    },
    {
        'STT': 3,
        'Nhom_Ton': 'Tồn năm 2025',
        'excel_row': 11,
        'ngay': '01/01/2026',
        'so_ct_phu': 'Số dư đầu kỳ',
        'so_ct_gl': '-',
        'so_hd': 'Tồn ĐK 2025',
        'dien_giai': 'Phần số dư Nợ đầu kỳ năm 2025 còn lại chưa có chứng từ kê khai giải trừ bên Có',
        'no': 324939777.0,
        'nguoi_lap': 'Kế toán EVN'
    }
])

print("df_ton_2025 sum:", df_ton_2025['no'].sum())
print("df_ton 2026 sum:", df_ton['no'].sum())
print("Total pending sum:", df_ton_2025['no'].sum() + df_ton['no'].sum())
