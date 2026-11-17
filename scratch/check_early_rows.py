import openpyxl

wb = openpyxl.load_workbook("GL_094_TK13895 (010126-311026).xlsx", data_only=True)
ws = wb.active

print("First 55 rows:")
for r in range(11, 55):
    f_val = ws.cell(row=r, column=6).value
    g_val = ws.cell(row=r, column=7).value
    e_val = ws.cell(row=r, column=5).value
    if f_val or g_val:
        print(f"Row {r}: Nợ={f_val} | Có={g_val} | {e_val}")
