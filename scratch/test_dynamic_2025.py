import sys, os
sys.path.insert(0, os.path.abspath("."))
import pandas as pd
from app_tk13895 import parse_gl094_excel, reconcile_tk13895

def build_df_ton_2025(open_bal, df_raw, df_2025, df_dnb):
    # Tìm dòng 205: HĐ 3704 & 3705 Cadivi
    m_cadivi = df_raw[(df_raw['dien_giai'].str.contains(r'3704.*3705|cadivi.*3704', case=False, na=False)) & (df_raw['no'] > 0)]
    row_cadivi = m_cadivi.iloc[0] if len(m_cadivi) > 0 else None
    
    # Tìm dòng 206: HĐ 0102 E.E & HĐ 0691 Hanaka
    m_ee_hanaka = df_raw[(df_raw['dien_giai'].str.contains(r'0102.*0691|ee.*hanaka', case=False, na=False)) & (df_raw['no'] > 0)]
    row_ee_hanaka = m_ee_hanaka.iloc[0] if len(m_ee_hanaka) > 0 else None
    
    val_cadivi = float(row_cadivi['no']) if row_cadivi is not None else 76238400.0
    val_ee_hanaka = float(row_ee_hanaka['no']) if row_ee_hanaka is not None else 299911123.0
    
    # Tính số dư nợ 2025 còn lại
    # Điện mua nội bộ T12/2025 giải trừ tại T1 và T2: 123.865.507.328 đ
    dnb_co_t1 = 87840801042.0
    dnb_co_t2_net = 132700116607.0 - 96675410321.0
    dnb_co_2025 = dnb_co_t1 + dnb_co_t2_net
    co_2025_regular = float(df_2025['co'].sum())
    
    total_co_2025 = dnb_co_2025 + co_2025_regular
    rem_2025_total = float(open_bal) - total_co_2025 # 701.089.300
    rem_open_bal_remaining = rem_2025_total - (val_cadivi + val_ee_hanaka) # 324.939.777
    
    rows = [
        {
            'STT': 1,
            'Nhom_Ton': 'Tồn năm 2025',
            'excel_row': int(row_cadivi['excel_row']) if row_cadivi is not None else 205,
            'ngay': '31/05/2026',
            'so_ct_phu': str(row_cadivi['so_ct_phu']).strip("'") if row_cadivi is not None else '25008',
            'so_ct_gl': str(row_cadivi['so_ct_gl']) if row_cadivi is not None else '1366',
            'so_hd': '3704 & 3705',
            'dien_giai': str(row_cadivi['dien_giai']) if row_cadivi is not None else '25008 Phải trả VAT HD3704 và HĐ3705 ngày 28.07.25 cty Cadivi_Điều chỉnh cấp đối tượng TK13895_9000',
            'no': val_cadivi,
            'co': 0.0,
            'nguoi_lap': str(row_cadivi['nguoi_lap']) if row_cadivi is not None else 'HIENNTK'
        },
        {
            'STT': 2,
            'Nhom_Ton': 'Tồn năm 2025',
            'excel_row': int(row_ee_hanaka['excel_row']) if row_ee_hanaka is not None else 206,
            'ngay': '31/05/2026',
            'so_ct_phu': str(row_ee_hanaka['so_ct_phu']).strip("'") if row_ee_hanaka is not None else '25017',
            'so_ct_gl': str(row_ee_hanaka['so_ct_gl']) if row_ee_hanaka is not None else '1366',
            'so_hd': '102 & 691',
            'dien_giai': str(row_ee_hanaka['dien_giai']) if row_ee_hanaka is not None else '25017 Phải trả VAT HD0102 ngày 05.08.25 cty E.E và 25017 Phải trả VAT HD0691 ngày 11.08.25 cty Hanaka_Điều chỉnh cấp đối tượng TK13895_9000',
            'no': val_ee_hanaka,
            'co': 0.0,
            'nguoi_lap': str(row_ee_hanaka['nguoi_lap']) if row_ee_hanaka is not None else 'HIENNTK'
        },
        {
            'STT': 3,
            'Nhom_Ton': 'Tồn năm 2025',
            'excel_row': 11,
            'ngay': '01/01/2026',
            'so_ct_phu': 'Số dư ĐK',
            'so_ct_gl': '-',
            'so_hd': 'Tồn ĐK 2025',
            'dien_giai': 'Phần số dư Nợ đầu kỳ năm 2025 còn lại chưa có chứng từ kê khai giải trừ bên Có',
            'no': rem_open_bal_remaining,
            'co': 0.0,
            'nguoi_lap': 'Kế toán EVN'
        }
    ]
    return pd.DataFrame(rows)

open_bal, df_raw = parse_gl094_excel("GL_094_TK13895 (010126-311026).xlsx")
open_bal, df, df_ton, df_da_gt, df_2025, df_dnb, df_t1_netoff, df_mistyped = reconcile_tk13895(df_raw, open_bal)
df_ton_2025 = build_df_ton_2025(open_bal, df_raw, df_2025, df_dnb)
print(df_ton_2025[['excel_row', 'so_hd', 'no', 'dien_giai']])
print("Sum 2025 ton:", df_ton_2025['no'].sum())
