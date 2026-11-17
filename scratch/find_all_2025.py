import openpyxl

wb = openpyxl.load_workbook("GL_094_TK13895 (010126-311026).xlsx", data_only=True)
ws = wb.active

print("Searching for 2025 or /25 or .25 in Dien Giai:")
rows_25 = []
for r in range(11, 597):
    dien_giai = str(ws.cell(row=r, column=5).value or '')
    no_val = ws.cell(row=r, column=6).value
    co_val = ws.cell(row=r, column=7).value
    ngay = ws.cell(row=r, column=2).value
    so_ct_phu = ws.cell(row=r, column=3).value
    so_ct_gl = ws.cell(row=r, column=4).value
    user = ws.cell(row=r, column=9).value
    
    # check 2025 patterns
    import re
    if re.search(r'2025|[/.-]25\b|041125|16\.12\.25|22\.12\.25|30/12/25', dien_giai, re.IGNORECASE):
        rows_25.append((r, ngay, so_ct_phu, so_ct_gl, dien_giai, no_val, co_val, user))

print(f"Total rows matched: {len(rows_25)}")
print("\n--- NỢ (Debit) rows ---")
for item in rows_25:
    if item[5] is not None:
        print(f"Row {item[0]}: Nợ={item[5]:,.0f} | Ngày={item[1]} | GL={item[3]} | Diễn giải: {item[4]}")

print("\n--- CÓ (Credit) rows ---")
for item in rows_25:
    if item[6] is not None:
        print(f"Row {item[0]}: Có={item[6]:,.0f} | Ngày={item[1]} | GL={item[3]} | Diễn giải: {item[4]}")
