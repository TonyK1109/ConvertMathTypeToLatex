from docxlatex import Document

document = Document('Test1.docx')
text = document.get_text().split('BẢNG ĐÁP ÁN VÀ HƯỚNG DẪN GIẢI')
print(text[1])
