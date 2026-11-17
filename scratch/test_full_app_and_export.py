import sys, os
sys.path.insert(0, os.path.abspath("."))
import pandas as pd
import openpyxl

from app_tk13895 import parse_gl094_excel, reconcile_tk13895, export_excel_report, build_df_ton_2025

print("Parsing GL_094 Excel...")
open_bal, df_raw = parse_gl094_excel("GL_094_TK13895 (010126-311026).xlsx")
print(f"Open bal: {open_bal:,.0f} đ, rows: {len(df_raw)}")

print("Running reconcile...")
open_bal, df, df_ton, df_da_gt, df_2025, df_dnb, df_t1_netoff, df_mistyped = reconcile_tk13895(df_raw, open_bal)
print(f"df_ton: {len(df_ton)} rows, sum: {df_ton['no'].sum():,.0f} đ")

print("Building df_ton_2025...")
df_ton_2025 = build_df_ton_2025(open_bal, df_raw, df_2025, df_dnb)
print(f"df_ton_2025: {len(df_ton_2025)} rows, sum: {df_ton_2025['no'].sum():,.0f} đ")
for _, r in df_ton_2025.iterrows():
    print(f"  Row {r['excel_row']} ({r['so_hd']}): {r['no']:,.0f} đ | {r['dien_giai'][:60]}...")

print("\nTesting export_excel_report...")
with open("GL_094_TK13895 (010126-311026).xlsx", "rb") as f:
    excel_stream = export_excel_report(
        open_bal=open_bal,
        df_ton=df_ton,
        df_da_gt=df_da_gt,
        df_2025=df_2025,
        df_dnb=df_dnb,
        df_t1_netoff=df_t1_netoff,
        df_mistyped=df_mistyped,
        uploaded_file=f,
        df_raw=df_raw
    )

test_out_path = "scratch/test_final_app_export.xlsx"
with open(test_out_path, "wb") as f_out:
    f_out.write(excel_stream.getvalue())
print(f"Saved: {test_out_path} ({len(excel_stream.getvalue()):,} bytes)")

wb = openpyxl.load_workbook(test_out_path, data_only=False)
print("Sheets in workbook:", wb.sheetnames)

# Verify Sheet 0
ws0 = wb["0. Sổ GL_094 Phân Loại"]
print(f"Sheet 0 max row: {ws0.max_row}, max col: {ws0.max_column}")
print(f"Sheet 0 Row 597: Col A={ws0['A597'].value}, Col J={ws0['J597'].value}, Col K={ws0['K597'].value}")

# Verify Sheet 1
ws1 = wb["1. HĐ Tồn Chưa Kê Khai"]
print(f"Sheet 1 title: {ws1['A1'].value}")
print(f"Sheet 1 Section I: {ws1['A4'].value}")
print(f"Sheet 1 Row 5: STT={ws1['A5'].value}, Nhóm={ws1['B5'].value}, Row={ws1['C5'].value}, HĐ={ws1['G5'].value}, Nợ={ws1['I5'].value}")
print(f"Sheet 1 Row 6: STT={ws1['A6'].value}, Nhóm={ws1['B6'].value}, Row={ws1['C6'].value}, HĐ={ws1['G6'].value}, Nợ={ws1['I6'].value}")
print(f"Sheet 1 Row 7: STT={ws1['A7'].value}, Nhóm={ws1['B7'].value}, Row={ws1['C7'].value}, HĐ={ws1['G7'].value}, Nợ={ws1['I7'].value}")
print(f"Sheet 1 Row 8 (Subtotal I): Label={ws1['A8'].value}, Formula={ws1['I8'].value}")
print(f"Sheet 1 Section II: {ws1['A9'].value}")
print(f"Sheet 1 Row 70 (Subtotal II): Label={ws1['A70'].value}, Formula={ws1['I70'].value}")
print(f"Sheet 1 Row 71 (Grand Total): Label={ws1['A71'].value}, Formula={ws1['I71'].value}")

print("\n--- ALL TESTS COMPLETED SUCCESSFULLY! ---")
