import sys, os
sys.path.insert(0, os.path.abspath("."))
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from app_tk13895 import parse_gl094_excel, reconcile_tk13895, build_df_ton_2025

open_bal, df_raw = parse_gl094_excel("GL_094_TK13895 (010126-311026).xlsx")
open_bal, df, df_ton, df_da_gt, df_2025, df_dnb, df_t1_netoff, df_mistyped = reconcile_tk13895(df_raw, open_bal)
df_ton_2025 = build_df_ton_2025(open_bal, df_raw, df_2025, df_dnb)

print("df_ton 2026 count:", len(df_ton), "sum:", df_ton['no'].sum())
print("df_ton_2025 count:", len(df_ton_2025), "sum:", df_ton_2025['no'].sum())

# Test creating 7 sheets
wb = openpyxl.load_workbook("GL_094_TK13895 (010126-311026).xlsx")
ws_main = wb.active
ws_main.title = "0. Sổ GL_094 Phân Loại"

font_title = Font(name="Times New Roman", size=14, bold=True, color="1E3A8A")
font_header = Font(name="Times New Roman", size=11, bold=True, color="FFFFFF")
font_bold = Font(name="Times New Roman", size=11, bold=True)
font_body = Font(name="Times New Roman", size=11)
fill_header = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
fill_total = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
fill_grand = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")
thin_border = Border(
    left=Side(style='thin', color='CBD5E1'),
    right=Side(style='thin', color='CBD5E1'),
    top=Side(style='thin', color='CBD5E1'),
    bottom=Side(style='thin', color='CBD5E1')
)

# 1. Sheet 1: HĐ 2026 Tồn Chưa Kê Khai
ws1 = wb.create_sheet(title="1. HĐ 2026 Tồn Chưa Kê Khai")
ws1.views.sheetView[0].showGridLines = True
ws1.merge_cells("A1:J1")
ws1["A1"] = "DANH SÁCH HÓA ĐƠN ĐẦU VÀO NĂM 2026 CÒN TỒN CHƯA KÊ KHAI (TK 13895)"
ws1["A1"].font = font_title
ws1["A1"].alignment = Alignment(vertical="center")
ws1.row_dimensions[1].height = 28

headers1 = ["STT", "Nhóm Tồn", "Dòng Excel", "Ngày CT", "Số CT Phụ", "Số CT GL", "Số HĐ Bóc Tách", "Diễn Giải Chi Tiết", "Số Phát Sinh Nợ", "Người Lập"]
ws1.append([])
ws1.append(headers1)
ws1.row_dimensions[3].height = 24
for c_idx in range(1, len(headers1) + 1):
    cell = ws1.cell(row=3, column=c_idx)
    cell.font = font_header
    cell.fill = fill_header
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = thin_border

cur_row = 4
stt = 1
for _, r in df_ton.iterrows():
    dt_str = r['ngay'].strftime('%d/%m/%Y') if hasattr(r['ngay'], 'strftime') else str(r['ngay'])[:10]
    ws1.append([
        stt,
        r['Nhom_Ton'],
        r['excel_row'],
        dt_str,
        str(r['so_ct_phu']).strip("'"),
        r['so_ct_gl'],
        r['so_hd'],
        r['dien_giai'],
        r['no'],
        r['nguoi_lap']
    ])
    for c in range(1, 11):
        cell = ws1.cell(row=cur_row, column=c)
        cell.font = font_body
        cell.border = thin_border
        if c in [1, 2, 3, 4, 5, 6, 7, 10]:
            cell.alignment = Alignment(horizontal="center", vertical="center")
        elif c == 9:
            cell.number_format = '#,##0'
            cell.alignment = Alignment(horizontal="right", vertical="center")
        else:
            cell.alignment = Alignment(horizontal="left", vertical="center")
    cur_row += 1
    stt += 1

ws1.merge_cells(start_row=cur_row, start_column=1, end_row=cur_row, end_column=8)
ws1.cell(row=cur_row, column=1).value = "TỔNG CỘNG HÓA ĐƠN NĂM 2026 CÒN TỒN"
ws1.cell(row=cur_row, column=1).font = font_bold
ws1.cell(row=cur_row, column=1).alignment = Alignment(horizontal="center", vertical="center")
cell_sum1 = ws1.cell(row=cur_row, column=9)
cell_sum1.value = f"=SUM(I4:I{cur_row-1})"
cell_sum1.font = font_bold
cell_sum1.number_format = '#,##0'
cell_sum1.alignment = Alignment(horizontal="right", vertical="center")
for c in range(1, 11):
    ws1.cell(row=cur_row, column=c).fill = fill_total
    ws1.cell(row=cur_row, column=c).border = thin_border
ws1.row_dimensions[cur_row].height = 22

widths1 = {1: 6, 2: 24, 3: 11, 4: 12, 5: 14, 6: 12, 7: 16, 8: 60, 9: 20, 10: 18}
for col_i, w in widths1.items():
    ws1.column_dimensions[get_column_letter(col_i)].width = w

# 2. Sheet 2: HĐ 2025 Tồn Chưa Kê Khai (SHEET RIÊNG)
ws2 = wb.create_sheet(title="2. HĐ 2025 Tồn Chưa Kê Khai")
ws2.views.sheetView[0].showGridLines = True
ws2.merge_cells("A1:J1")
ws2["A1"] = "DANH SÁCH HÓA ĐƠN & SỐ DƯ ĐẦU KỲ NĂM 2025 CÒN TỒN CHƯA KÊ KHAI (TK 13895)"
ws2["A1"].font = font_title
ws2["A1"].alignment = Alignment(vertical="center")
ws2.row_dimensions[1].height = 28

ws2.merge_cells("A2:J2")
ws2["A2"] = "Số dư nợ đầu kỳ: 125.834.024.708 đ | Đã kê khai giải trừ trong kỳ 2026: 125.132.935.408 đ | Số dư tồn chưa kê khai: 701.089.300 đ"
ws2["A2"].font = Font(name="Times New Roman", size=10, italic=True, color="475569")
ws2["A2"].alignment = Alignment(vertical="center")
ws2.row_dimensions[2].height = 20

headers2 = ["STT", "Phân Loại Tồn", "Dòng Excel", "Ngày CT", "Số CT Phụ", "Số CT GL", "Số HĐ Bóc Tách", "Diễn Giải Chi Tiết", "Số Tiền Còn Tồn", "Người Lập / Ghi Chú"]
ws2.append([])
ws2.append(headers2)
ws2.row_dimensions[4].height = 24
for c_idx in range(1, len(headers2) + 1):
    cell = ws2.cell(row=4, column=c_idx)
    cell.font = font_header
    cell.fill = fill_header
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = thin_border

cur_row2 = 5
for _, r25 in df_ton_2025.iterrows():
    dt_str = r25['ngay'].strftime('%d/%m/%Y') if hasattr(r25['ngay'], 'strftime') else str(r25['ngay'])[:10]
    ws2.append([
        r25['STT'],
        r25['Nhom_Ton'],
        r25['excel_row'],
        dt_str,
        str(r25['so_ct_phu']).strip("'"),
        r25['so_ct_gl'],
        r25['so_hd'],
        r25['dien_giai'],
        r25['no'],
        r25['nguoi_lap']
    ])
    for c in range(1, 11):
        cell = ws2.cell(row=cur_row2, column=c)
        cell.font = font_body
        cell.border = thin_border
        if c in [1, 2, 3, 4, 5, 6, 7]:
            cell.alignment = Alignment(horizontal="center", vertical="center")
        elif c == 9:
            cell.number_format = '#,##0'
            cell.alignment = Alignment(horizontal="right", vertical="center")
        else:
            cell.alignment = Alignment(horizontal="left", vertical="center")
    cur_row2 += 1

ws2.merge_cells(start_row=cur_row2, start_column=1, end_row=cur_row2, end_column=8)
ws2.cell(row=cur_row2, column=1).value = "TỔNG CỘNG HÓA ĐƠN & SỐ DƯ TỒN NĂM 2025 CHƯA KÊ KHAI"
ws2.cell(row=cur_row2, column=1).font = font_bold
ws2.cell(row=cur_row2, column=1).alignment = Alignment(horizontal="center", vertical="center")
cell_sum2 = ws2.cell(row=cur_row2, column=9)
cell_sum2.value = f"=SUM(I5:I{cur_row2-1})"
cell_sum2.font = Font(name="Times New Roman", size=12, bold=True, color="991B1B")
cell_sum2.number_format = '#,##0'
cell_sum2.alignment = Alignment(horizontal="right", vertical="center")
for c in range(1, 11):
    ws2.cell(row=cur_row2, column=c).fill = fill_grand
    ws2.cell(row=cur_row2, column=c).border = thin_border
ws2.row_dimensions[cur_row2].height = 25

widths2 = {1: 6, 2: 18, 3: 11, 4: 12, 5: 14, 6: 12, 7: 16, 8: 65, 9: 20, 10: 20}
for col_i, w in widths2.items():
    ws2.column_dimensions[get_column_letter(col_i)].width = w

print("Sheets in wb:", wb.sheetnames)
wb.save("scratch/test_7_sheets_wb.xlsx")
print("Saved 7-sheet test workbook!")
