from docxlatex import Document
import json

suject = "Toán"
data = {
    "subject": suject,
    "body": "Tính thể tích của khối lập phương có cạnh bằng 2",
    "answers": [
        {"value": "6"},
        {"value": "4"},
        {"value": "2"},
        {"value": "8", "isCorrect": True}
    ],
    "level": 1
}

datas = []
datas.append(data)
with open('data.json', 'w', encoding='utf-8') as f:
    json.dump(datas, f, ensure_ascii=False,  indent=4)

document = Document('Test1.docx')
text = document.get_text().split('BẢNG ĐÁP ÁN VÀ HƯỚNG DẪN GIẢI')
for x in text[1].split("Lời giải"):
    print(x, 'endline \n')
