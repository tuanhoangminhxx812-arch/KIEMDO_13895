import docx

doc = docx.Document("Ứng dụng kiểm dò 13895 (thêm).docx")
for i in range(30, 46):
    if i < len(doc.paragraphs):
        print(f"[{i}] {doc.paragraphs[i].text.strip()}")
