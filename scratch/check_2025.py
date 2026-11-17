import docx

for fn in ['Ứng dụng kiểm dò 13895.docx', 'Ứng dụng kiểm dò 13895 (thêm).docx']:
    try:
        doc = docx.Document(fn)
        print('=== ' + fn + ' ===')
        for i, p in enumerate(doc.paragraphs):
            t = p.text.strip()
            if any(k in t.lower() for k in ['2025', 'tồn', 'đầu kỳ', 'cadivi', 'hanaka', '0102', '3704']):
                print(f'P{i}: {t}')
        for i, tbl in enumerate(doc.tables):
            for r_idx, row in enumerate(tbl.rows):
                row_txt = " | ".join(c.text.strip() for c in row.cells)
                if any(k in row_txt.lower() for k in ['2025', 'tồn', 'đầu kỳ', 'cadivi', 'hanaka', '0102', '3704']):
                    print(f'Table {i} Row {r_idx}: {row_txt}')
    except Exception as e:
        print(f'Error reading {fn}: {e}')
