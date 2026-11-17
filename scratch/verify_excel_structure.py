import openpyxl

wb = openpyxl.load_workbook("scratch/test_final_app_export.xlsx")

print("Checking Sheet 1 in test_final_app_export.xlsx:")
ws1 = wb["1. HĐ Tồn Chưa Kê Khai"]

print(f"Max row: {ws1.max_row}, Max col: {ws1.max_column}")
for r in [1, 3, 4, 5, 6, 7, 8, 9, 10, 69, 70, 71]:
    vals = [ws1.cell(row=r, column=c).value for c in range(1, 11)]
    print(f"Row {r:2d}: {vals}")

# Check Sheet 0
ws0 = wb["0. Sổ GL_094 Phân Loại"]
print("\nChecking Sheet 0 in test_final_app_export.xlsx:")
for r in [11, 204, 205, 206, 595, 596, 597]:
    j_val = ws0.cell(row=r, column=10).value
    k_val = ws0.cell(row=r, column=11).value
    l_val = ws0.cell(row=r, column=12).value
    print(f"Row {r:3d}: Col J='{j_val}' | Col K={k_val} | Col L='{str(l_val)[:40]}...'")
