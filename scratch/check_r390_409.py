import openpyxl

wb = openpyxl.load_workbook("GL_094_TK13895 (010126-311026).xlsx", data_only=True)
ws = wb.active

for r in [388, 389, 390, 409]:
    f_val = ws.cell(row=r, column=6).value
    g_val = ws.cell(row=r, column=7).value
    e_val = ws.cell(row=r, column=5).value
    c_val = ws.cell(row=r, column=3).value
    d_val = ws.cell(row=r, column=4).value
    print(f"Row {r}: Nợ={f_val} | Có={g_val} | AP={c_val} | GL={d_val} | Diễn giải={e_val}")
