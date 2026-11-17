import openpyxl

wb = openpyxl.load_workbook("GL_094_TK13895 (010126-311026).xlsx", data_only=True)
ws = wb.active

print("Inspecting Rows 200 to 215:")
for r in range(200, 215):
    vals = [ws.cell(row=r, column=c).value for c in range(1, 11)]
    print(f"Row {r}: col B={vals[1]} | col C={vals[2]} | col D={vals[3]} | col E={vals[4]} | col F={vals[5]} | col G={vals[6]} | col H={vals[7]}")
