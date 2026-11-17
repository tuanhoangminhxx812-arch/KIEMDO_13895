import docx

doc = docx.Document("Ứng dụng kiểm dò 13895 (thêm).docx")
for i, p in enumerate(doc.paragraphs):
    if p.text.strip():
        print(f"[{i}] {p.text.strip()}")
