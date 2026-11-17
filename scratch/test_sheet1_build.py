import io, sys, os
sys.path.insert(0, os.path.abspath("."))
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from app_tk13895 import parse_gl094_excel, reconcile_tk13895

open_bal, df_raw = parse_gl094_excel("GL_094_TK13895 (010126-311026).xlsx")
open_bal, df, df_ton, df_da_gt, df_2025, df_dnb, df_t1_netoff, df_mistyped = reconcile_tk13895(df_raw, open_bal)

wb = openpyxl.Workbook()
ws1 = wb.active
ws1.title = "1. HĐ Tồn Chưa Kê Khai"
ws1.views.sheetView[0].showGridLines = True

font_title = Font(name="Times New Roman", size=14, bold=True, color="1E3A8A")
font_section = Font(name="Times New Roman", size=11, bold=True, color="1E3A8A")
font_header = Font(name="Times New Roman", size=11, bold=True, color="FFFFFF")
font_subtotal = Font(name="Times New Roman", size=11, bold=True, color="0F172A")
font_bold = Font(name="Times New Roman", size=11, bold=True)
font_body = Font(name="Times New Roman", size=11)

fill_header = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
fill_section = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")
fill_subtotal = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
fill_grand_total = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")

thin_border = Border(
    left=Side(style='thin', color='CBD5E1'),
    right=Side(style='thin', color='CBD5E1'),
    top=Side(style='thin', color='CBD5E1'),
    bottom=Side(style='thin', color='CBD5E1')
)
double_bottom_border = Border(
    left=Side(style='thin', color='CBD5E1'),
    right=Side(style='thin', color='CBD5E1'),
    top=Side(style='thin', color='CBD5E1'),
    bottom=Side(style='double', color='0F172A')
)

ws1.merge_cells("A1:J1")
ws1["A1"] = "DANH SÁCH HÓA ĐƠN ĐẦU VÀO CÒN TỒN CHƯA KÊ KHAI (NĂM 2025 & NĂM 2026) - TK 13895"
ws1["A1"].font = font_title
ws1["A1"].alignment = Alignment(vertical="center")
ws1.row_dimensions[1].height = 28

headers1 = ["STT", "Năm / Nhóm Tồn", "Dòng Excel", "Ngày CT", "Số CT Phụ", "Số CT GL", "Số HĐ Bóc Tách", "Diễn Giải Chi Tiết", "Số Phát Sinh Nợ", "Người Lập"]
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

# --- PHẦN I: NĂM 2025 ---
ws1.merge_cells(start_row=cur_row, start_column=1, end_row=cur_row, end_column=10)
sec1_cell = ws1.cell(row=cur_row, column=1)
sec1_cell.value = "I. HÓA ĐƠN & SỐ DƯ ĐẦU KỲ NĂM 2025 CÒN TỒN CHƯA KÊ KHAI GIẢI TRỪ (701.089.300 đ)"
sec1_cell.font = font_section
sec1_cell.fill = fill_section
sec1_cell.alignment = Alignment(horizontal="left", vertical="center")
ws1.row_dimensions[cur_row].height = 22
for col_idx in range(1, 11):
    ws1.cell(row=cur_row, column=col_idx).border = thin_border
cur_row += 1

r_start_2025 = cur_row
df_ton_2025_data = [
    (1, "Tồn năm 2025", 205, "31/05/2026", "25008", "1366", "3704 & 3705", "25008 Phải trả VAT HD3704 và HĐ3705 ngày 28.07.25 cty Cadivi_Điều chỉnh cấp đối tượng TK13895_9000", 76238400.0, "HIENNTK"),
    (2, "Tồn năm 2025", 206, "31/05/2026", "25017", "1366", "102 & 691", "25017 Phải trả VAT HD0102 ngày 05.08.25 cty E.E và 25017 Phải trả VAT HD0691 ngày 11.08.25 cty Hanaka_Điều chỉnh cấp đối tượng TK13895_9000", 299911123.0, "HIENNTK"),
    (3, "Tồn năm 2025", 11, "01/01/2026", "Số dư đầu kỳ", "-", "Tồn ĐK 2025", "Phần số dư Nợ đầu kỳ năm 2025 còn lại chưa có chứng từ kê khai giải trừ bên Có", 324939777.0, "Kế toán EVN")
]

for row_data in df_ton_2025_data:
    ws1.append(list(row_data))
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

r_end_2025 = cur_row - 1

# Subtotal Phần I
ws1.merge_cells(start_row=cur_row, start_column=1, end_row=cur_row, end_column=8)
sub1_title = ws1.cell(row=cur_row, column=1)
sub1_title.value = "CỘNG PHẦN I: HÓA ĐƠN & SỐ DƯ TỒN NĂM 2025"
sub1_title.font = font_subtotal
sub1_title.alignment = Alignment(horizontal="center", vertical="center")

sub1_val = ws1.cell(row=cur_row, column=9)
sub1_val.value = f"=SUM(I{r_start_2025}:I{r_end_2025})"
sub1_val.font = font_subtotal
sub1_val.number_format = '#,##0'
sub1_val.alignment = Alignment(horizontal="right", vertical="center")

for c in range(1, 11):
    c_cell = ws1.cell(row=cur_row, column=c)
    c_cell.fill = fill_subtotal
    c_cell.border = thin_border

r_sub_2025 = cur_row
cur_row += 1

# --- PHẦN II: NĂM 2026 ---
ws1.merge_cells(start_row=cur_row, start_column=1, end_row=cur_row, end_column=10)
sec2_cell = ws1.cell(row=cur_row, column=1)
sec2_cell.value = "II. HÓA ĐƠN PHÁT SINH NĂM 2026 CÒN TỒN CHƯA KÊ KHAI GIẢI TRỪ (1.995.027.119 đ)"
sec2_cell.font = font_section
sec2_cell.fill = fill_section
sec2_cell.alignment = Alignment(horizontal="left", vertical="center")
ws1.row_dimensions[cur_row].height = 22
for col_idx in range(1, 11):
    ws1.cell(row=cur_row, column=col_idx).border = thin_border
cur_row += 1

r_start_2026 = cur_row
stt_2026 = 1
for _, r in df_ton.iterrows():
    dt_str = r['ngay'].strftime('%d/%m/%Y') if hasattr(r['ngay'], 'strftime') else str(r['ngay'])[:10]
    ws1.append([
        stt_2026,
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
    stt_2026 += 1

r_end_2026 = cur_row - 1

# Subtotal Phần II
ws1.merge_cells(start_row=cur_row, start_column=1, end_row=cur_row, end_column=8)
sub2_title = ws1.cell(row=cur_row, column=1)
sub2_title.value = "CỘNG PHẦN II: HÓA ĐƠN TỒN NĂM 2026 (60 HÓA ĐƠN)"
sub2_title.font = font_subtotal
sub2_title.alignment = Alignment(horizontal="center", vertical="center")

sub2_val = ws1.cell(row=cur_row, column=9)
sub2_val.value = f"=SUM(I{r_start_2026}:I{r_end_2026})"
sub2_val.font = font_subtotal
sub2_val.number_format = '#,##0'
sub2_val.alignment = Alignment(horizontal="right", vertical="center")

for c in range(1, 11):
    c_cell = ws1.cell(row=cur_row, column=c)
    c_cell.fill = fill_subtotal
    c_cell.border = thin_border

r_sub_2026 = cur_row
cur_row += 1

# --- GRAND TOTAL ROW ---
ws1.merge_cells(start_row=cur_row, start_column=1, end_row=cur_row, end_column=8)
tot_title = ws1.cell(row=cur_row, column=1)
tot_title.value = "TỔNG CỘNG HÓA ĐƠN CÒN TỒN CHƯA KÊ KHAI (NĂM 2025 & NĂM 2026)"
tot_title.font = font_bold
tot_title.alignment = Alignment(horizontal="center", vertical="center")

tot_val = ws1.cell(row=cur_row, column=9)
tot_val.value = f"=I{r_sub_2025}+I{r_sub_2026}"
tot_val.font = Font(name="Times New Roman", size=12, bold=True, color="991B1B")
tot_val.number_format = '#,##0'
tot_val.alignment = Alignment(horizontal="right", vertical="center")

for c in range(1, 11):
    c_cell = ws1.cell(row=cur_row, column=c)
    c_cell.fill = fill_grand_total
    c_cell.border = double_bottom_border
ws1.row_dimensions[cur_row].height = 25

widths1 = {1: 6, 2: 24, 3: 12, 4: 12, 5: 14, 6: 12, 7: 15, 8: 50, 9: 18, 10: 16}
for col_i, w in widths1.items():
    ws1.column_dimensions[get_column_letter(col_i)].width = w

output_test = "scratch/test_sheet1_2025_2026.xlsx"
wb.save(output_test)
print(f"Saved test file: {output_test}")
print(f"r_sub_2025={r_sub_2025}, formula={sub1_val.value}")
print(f"r_sub_2026={r_sub_2026}, formula={sub2_val.value}")
print(f"r_grand_total={cur_row}, formula={tot_val.value}")
