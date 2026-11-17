import openpyxl

wb = openpyxl.load_workbook("GL_094_TK13895 (010126-311026).xlsx", data_only=True)
ws = wb.active

print("GL 1366 rows:")
for r in range(11, 597):
    gl = ws.cell(row=r, column=4).value
    if str(gl) == '1366':
        print(f"Row {r}: Nợ={ws.cell(row=r, column=6).value} | Có={ws.cell(row=r, column=7).value} | {ws.cell(row=r, column=5).value}")
