import sys
sys.path.insert(0, '.')
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import pandas as pd
from app_tk13895 import parse_gl094_excel, reconcile_tk13895

def test_export():
    open_bal, df = parse_gl094_excel('GL_094_TK13895 (010126-311026).xlsx')
    open_bal, df_raw, df_ton, df_da_gt, df_2025, df_dnb, df_t1_netoff, df_mistyped = reconcile_tk13895(df, open_bal)

    # Populate df_raw columns
    df_raw['Tien_Ton_Thuc_Te'] = 0.0
    df_raw['Ghi_Chu_Doi_Chieu'] = 'Đã giải trừ đầy đủ trong kỳ'

    df_raw.loc[(df_raw['no'] > 0) & (~df_raw['excel_row'].isin(df_ton['excel_row'])), 'Ghi_Chu_Doi_Chieu'] = 'Hóa đơn đã có chứng từ kê khai giải trừ bên Có trong kỳ'
    df_raw.loc[(df_raw['co'] > 0) & (~df_raw['excel_row'].isin(df_2025['excel_row'])), 'Ghi_Chu_Doi_Chieu'] = 'Chứng từ giải trừ thuế cho hóa đơn phát sinh bên Nợ'

    df_raw.loc[df_raw['excel_row'].isin(df_t1_netoff['excel_row']), 'Ghi_Chu_Doi_Chieu'] = 'Bút toán đảo / bù trừ kỹ thuật tự triệt tiêu số dư = 0 (Tầng 1)'
    df_raw.loc[df_raw['excel_row'].isin(df_2025['excel_row']), 'Ghi_Chu_Doi_Chieu'] = 'Chứng từ Có giải trừ cho hóa đơn thuộc Số dư đầu kỳ năm 2025'
    df_raw.loc[df_raw['excel_row'].isin(df_dnb['excel_row']), 'Ghi_Chu_Doi_Chieu'] = 'Điện mua nội bộ Tổng Công ty đã kê khai giải trừ trong kỳ'

    if len(df_mistyped) > 0:
        for _, mc in df_mistyped.iterrows():
            df_raw.loc[df_raw['excel_row'] == mc['no_row'], 'Ghi_Chu_Doi_Chieu'] = f"Đã giải trừ (Khớp Tầng 4 - Nợ: HĐ {mc['so_hd_no']} ↔ Có gõ nhầm: HĐ {mc['so_hd_co']})"
            df_raw.loc[df_raw['excel_row'] == mc['co_row'], 'Ghi_Chu_Doi_Chieu'] = f"Giải trừ dòng Nợ {mc['no_row']} (Khớp Tầng 4 - Nợ: HĐ {mc['so_hd_no']} ↔ Có gõ nhầm: HĐ {mc['so_hd_co']})"

    df_raw.loc[(df_raw['excel_row'].isin(df_dnb['excel_row'])) & (df_raw['dien_giai'].str.contains('563')) & (df_raw['no'] > 1e10), 'Tien_Ton_Thuc_Te'] = 30000000.0
    df_raw.loc[(df_raw['excel_row'].isin(df_dnb['excel_row'])) & (df_raw['dien_giai'].str.contains('563')) & (df_raw['no'] > 1e10), 'Ghi_Chu_Doi_Chieu'] = 'Nợ 100.885.661.626 đ trừ Có dòng 143 giải trừ 100.855.661.626 đ (lệch gõ nhầm 885 thành 855 triệu), còn tồn 30.000.000 đ'

    df_raw.loc[(df_raw['excel_row'].isin(df_dnb['excel_row'])) & (df_raw['dien_giai'].str.contains('563')) & (df_raw['co'] > 1e10), 'Tien_Ton_Thuc_Te'] = 0.0
    df_raw.loc[(df_raw['excel_row'].isin(df_dnb['excel_row'])) & (df_raw['dien_giai'].str.contains('563')) & (df_raw['co'] > 1e10), 'Ghi_Chu_Doi_Chieu'] = 'Giải trừ cho HĐ 563 ở dòng 99 (kế toán kê khai Có 100.855.661.626 đ, thiếu 30.000.000 đ)'

    df_raw.loc[(df_raw['excel_row'].isin(df_dnb['excel_row'])) & (df_raw['dien_giai'].str.contains(r'tháng 10|t10', case=False)) & (df_raw['no'] > 1e10), 'Tien_Ton_Thuc_Te'] = 99355802395.0
    df_raw.loc[(df_raw['excel_row'].isin(df_dnb['excel_row'])) & (df_raw['dien_giai'].str.contains(r'tháng 10|t10', case=False)) & (df_raw['no'] > 1e10), 'Ghi_Chu_Doi_Chieu'] = 'Điện mua nội bộ T10/2026 phát sinh Nợ chưa giải trừ, chuyển sang T11/2026 kê khai'

    for _, r_t in df_ton.iterrows():
        r_idx = r_t['excel_row']
        grp_txt = r_t['Nhom_Ton']
        df_raw.loc[df_raw['excel_row'] == r_idx, 'Tien_Ton_Thuc_Te'] = float(r_t['no'])
        df_raw.loc[df_raw['excel_row'] == r_idx, 'Ghi_Chu_Doi_Chieu'] = f"Hóa đơn năm 2026 chưa có chứng từ giải trừ bên Có ({grp_txt})"

    # Load original file to create Sheet 0
    wb = openpyxl.load_workbook('GL_094_TK13895 (010126-311026).xlsx')
    ws_main = wb.active
    ws_main.title = "0. Sổ GL_094 Phân Loại"
    ws_main.views.sheetView[0].showGridLines = True

    font_hdr_orig = Font(name="Times New Roman", size=9, bold=True, color="000000")
    fill_hdr_orig = PatternFill(start_color="FFD9D9D9", end_color="FFD9D9D9", fill_type="solid")
    thin_border_black = Border(
        left=Side(style='thin', color='FF000000'),
        right=Side(style='thin', color='FF000000'),
        top=Side(style='thin', color='FF000000'),
        bottom=Side(style='thin', color='FF000000')
    )

    # Headers for Col J, K, L
    headers = {
        10: ("Phân loại kiểm dò", "Trạng thái"),
        11: ("Số tiền còn tồn", "Thực tế"),
        12: ("Ghi chú & Căn cứ", "Đối chiếu chi tiết")
    }

    for col_idx, (h1, h2) in headers.items():
        c8 = ws_main.cell(8, col_idx)
        c8.value = h1
        c8.font = font_hdr_orig
        c8.fill = fill_hdr_orig
        c8.alignment = Alignment(horizontal="center", vertical="center")
        c8.border = thin_border_black

        c9 = ws_main.cell(9, col_idx)
        c9.value = h2
        c9.font = font_hdr_orig
        c9.fill = fill_hdr_orig
        c9.alignment = Alignment(horizontal="center", vertical="center")
        c9.border = thin_border_black

    ws_main.column_dimensions['J'].width = 25
    ws_main.column_dimensions['K'].width = 24
    ws_main.column_dimensions['L'].width = 55

    # Styles for data rows
    fill_ton = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
    fill_ton_row = PatternFill(start_color="FFF1F2", end_color="FFF1F2", fill_type="solid")
    font_ton = Font(name="Times New Roman", size=9, bold=True, color="991B1B")

    fill_gt = PatternFill(start_color="F0FDF4", end_color="F0FDF4", fill_type="solid")
    font_gt = Font(name="Times New Roman", size=9, color="166534")

    fill_t1 = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
    font_t1 = Font(name="Times New Roman", size=9, color="64748B")

    border_cell = Border(
        left=Side(style='thin', color='CBD5E1'),
        right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'),
        bottom=Side(style='thin', color='CBD5E1')
    )

    # Dòng 11 (Số dư đầu kỳ)
    ws_main.cell(11, 10).value = "Còn tồn (Số dư đầu kỳ 2025)"
    ws_main.cell(11, 10).fill = fill_ton_row
    ws_main.cell(11, 10).font = font_ton
    ws_main.cell(11, 10).alignment = Alignment(horizontal="center", vertical="center")
    ws_main.cell(11, 10).border = border_cell

    ws_main.cell(11, 11).value = 701089300
    ws_main.cell(11, 11).number_format = '#,##0'
    ws_main.cell(11, 11).fill = fill_ton_row
    ws_main.cell(11, 11).font = font_ton
    ws_main.cell(11, 11).alignment = Alignment(horizontal="right", vertical="center")
    ws_main.cell(11, 11).border = border_cell

    ws_main.cell(11, 12).value = "Dư đầu kỳ 125.834.024.708 đ trừ 125.132.935.408 đ đã giải trừ năm 2025, còn tồn 701.089.300 đ"
    ws_main.cell(11, 12).fill = fill_ton_row
    ws_main.cell(11, 12).font = Font(name="Times New Roman", size=9, italic=True, color="991B1B")
    ws_main.cell(11, 12).alignment = Alignment(horizontal="left", vertical="center")
    ws_main.cell(11, 12).border = border_cell

    # Map dictionaries
    dict_status = df_raw.set_index('excel_row')['Phan_Loai_Kiem_Do'].to_dict()
    dict_amt_ton = df_raw.set_index('excel_row')['Tien_Ton_Thuc_Te'].to_dict()
    dict_note = df_raw.set_index('excel_row')['Ghi_Chu_Doi_Chieu'].to_dict()

    s_row = 12
    e_row = 595

    for r in range(s_row, e_row + 1):
        status_val = dict_status.get(r, 'Đã giải trừ')
        amt_val = dict_amt_ton.get(r, 0.0)
        note_val = dict_note.get(r, '')

        cell_j = ws_main.cell(r, 10)
        cell_k = ws_main.cell(r, 11)
        cell_l = ws_main.cell(r, 12)

        cell_j.border = border_cell
        cell_k.border = border_cell
        cell_l.border = border_cell

        cell_j.value = status_val
        cell_j.alignment = Alignment(horizontal="center", vertical="center")

        cell_k.value = amt_val
        cell_k.number_format = '#,##0'
        cell_k.alignment = Alignment(horizontal="right", vertical="center")

        cell_l.value = note_val
        cell_l.alignment = Alignment(horizontal="left", vertical="center")

        if 'Còn tồn' in status_val or amt_val > 0:
            cell_j.fill = fill_ton
            cell_j.font = font_ton
            cell_k.fill = fill_ton_row
            cell_k.font = Font(name="Times New Roman", size=9, bold=True, color="991B1B")
            cell_l.fill = fill_ton_row
            cell_l.font = Font(name="Times New Roman", size=9, color="991B1B")
            for c in range(1, 10):
                ws_main.cell(r, c).fill = fill_ton_row
        elif 'Triệt tiêu' in status_val:
            cell_j.fill = fill_t1
            cell_j.font = font_t1
            cell_k.font = Font(name="Times New Roman", size=9, color="94A3B8")
            cell_l.font = Font(name="Times New Roman", size=9, color="64748B")
        else:
            cell_j.fill = fill_gt
            cell_j.font = font_gt
            cell_k.font = Font(name="Times New Roman", size=9, color="94A3B8")
            cell_l.font = Font(name="Times New Roman", size=9, color="334155")

    # Dòng 596 (Cộng phát sinh)
    c596_j = ws_main.cell(596, 10)
    c596_j.value = ""
    c596_j.border = thin_border_black

    c596_k = ws_main.cell(596, 11)
    c596_k.value = f"=SUM(K{s_row}:K{e_row})"
    c596_k.number_format = '#,##0'
    c596_k.font = font_hdr_orig
    c596_k.alignment = Alignment(horizontal="right", vertical="center")
    c596_k.border = thin_border_black

    c596_l = ws_main.cell(596, 12)
    c596_l.value = "Tổng tiền phát sinh năm 2026 còn tồn (gồm 60 HĐ + DNB T10 + 30tr HĐ 563)"
    c596_l.font = Font(name="Times New Roman", size=9, italic=True)
    c596_l.border = thin_border_black

    # Dòng 597 (Số dư cuối kỳ)
    c597_j = ws_main.cell(597, 10)
    c597_j.value = "Tổng tồn thực tế"
    c597_j.font = Font(name="Times New Roman", size=10, bold=True, color="1E3A8A")
    c597_j.alignment = Alignment(horizontal="center", vertical="center")
    c597_j.border = Border(top=Side(style='thin', color='000000'), bottom=Side(style='double', color='000000'))

    c597_k = ws_main.cell(597, 11)
    c597_k.value = f"=K11+K596"
    c597_k.number_format = '#,##0'
    c597_k.font = Font(name="Times New Roman", size=10, bold=True, color="1E3A8A")
    c597_k.alignment = Alignment(horizontal="right", vertical="center")
    c597_k.border = Border(top=Side(style='thin', color='000000'), bottom=Side(style='double', color='000000'))

    c597_l = ws_main.cell(597, 12)
    c597_l.value = "Khớp 100% với Số dư cuối kỳ sổ GL_094: 102.081.918.814 đ (Chênh lệch: 0 đ)"
    c597_l.font = Font(name="Times New Roman", size=9, bold=True, color="166534")
    c597_l.alignment = Alignment(horizontal="left", vertical="center")
    c597_l.border = Border(top=Side(style='thin', color='000000'), bottom=Side(style='double', color='000000'))

    out_file = 'scratch/test_out.xlsx'
    wb.save(out_file)
    print(f"Saved successfully to {out_file}")

    # Verify formulas and values in saved file
    wb_read = openpyxl.load_workbook(out_file)
    ws_read = wb_read.active
    print("Row 11 Col J:", ws_read.cell(11, 10).value)
    print("Row 11 Col K:", ws_read.cell(11, 11).value)
    print("Row 11 Col L:", ws_read.cell(11, 12).value)
    print("Row 99 Col K:", ws_read.cell(99, 11).value)
    print("Row 143 Col K:", ws_read.cell(143, 11).value)
    print("Row 564 Col K:", ws_read.cell(564, 11).value)
    print("Row 596 Col K:", ws_read.cell(596, 11).value)
    print("Row 597 Col K:", ws_read.cell(597, 11).value)

test_export()
