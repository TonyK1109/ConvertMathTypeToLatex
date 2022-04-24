# from docxlatex import Document
# import json

# suject = "Toán"
# data = {
#     "subject": suject,
#     "body": "Tính thể tích của khối lập phương có cạnh bằng 2",
#     "answers": [
#         {"value": "6"},
#         {"value": "4"},
#         {"value": "2"},
#         {"value": "8", "isCorrect": True}
#     ],
#     "level": 1
# }

# datas = []
# datas.append(data)
# with open('data.json', 'w', encoding='utf-8') as f:
#     json.dump(datas, f, ensure_ascii=False,  indent=4)

# document = Document('Test1.docx')
# text = document.get_text().split('BẢNG ĐÁP ÁN VÀ HƯỚNG DẪN GIẢI')
# for x in text[1].split("Lời giải"):
#     print(x, 'endline \n')
import zipfile
import re
import os

from defusedxml import ElementTree
from dwmlLocal import omml

extensions = ['.jpg', '.jpeg', '.png', '.svg', '.bmp', '.gif']
inline_delimiter = '$'
ns_map = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
}

def qn(tag):
    """
    A utility function to turn a namespace
    prefixed tag name into a Clark-notation qualified tag name for lxml. For
    example, qn('m:oMath') returns '{http://schemas.openxmlformats.org/officeDocument/2006/math}oMath'

    :param tag:str - A namespace-prefixed tag name
    :return qn:str - A Clark-notation qualified name tag for lxml.
    """
    prefix, tag_root = tag.split(':')
    uri = ns_map[prefix]
    return '{{{}}}{}'.format(uri, tag_root)

def tag_to_latex(tag):
    xmlstr = ElementTree.tostring(tag, encoding='unicode', method='xml')
    xmlStr = '''<MathEquation>''' + xmlstr + '''</MathEquation>'''
    for omath in omml.load_string(xmlStr):
        return omath.latex


zip_f = zipfile.ZipFile('Original.docx')
print(zip_f.namelist())

for f in zip_f.namelist():
    if f.startswith('word/document'):
        t = zip_f.read(f)
    if f == 'word/_rels/document.xml.rels':
        im = zip_f.read(f)

# for f in zip_f.namelist():
#     _, extension = os.path.splitext(f)
#     if extension in extensions:
#         destination = os.path.join("image", os.path.basename(f))
#         with open(destination, 'wb') as destination_file:
#             destination_file.write(zip_f.read(f))

zip_f.close()


# with open('original.xml', 'wb') as fa:
#     fa.write(t)


# with open('test.rels.xml', 'wb') as fa:
#     fa.write(im)
equations = []
text = ''
n_images = 0
root = ElementTree.fromstring(t)
for child in root.iter():

    if child.tag == qn('w:bookmarkStart'):
        text += '&'

    if child.tag == qn('w:t'):
        text += child.text if child.text is not None else ''

    # Found an equation
    elif child.tag == qn('m:oMath'):
        text += inline_delimiter + ' '
        equations.append(child)
        text += tag_to_latex(child)
        text += ' ' + inline_delimiter
    # elif child.tag == qn('m:r'):
    #     text += ''.join(child.itertext())

    # Found an image
    elif child.tag == qn('w:drawing'):
        n_images += 1
        text += f'\nIMAGE#{n_images}-image{n_images}\n'

    elif child.tag == qn('w:tab'):
        text += '\t'
    elif child.tag == qn('w:br') or child.tag == qn('w:cr'):
        text += '\n'
    elif child.tag == qn('w:p'):
        text += '\n\n'

text = re.sub(r'\n(\n+)\$(\s*.+\s*)\$\n', r'\n\1$$ \2 $$', text)

print(text.split('BẢNG ĐÁP ÁN VÀ HƯỚNG DẪN GIẢI')[1])
# print(text.split('&')[3].split('\n')[0])
# tree = ET.parse('tests/composite.xml')

# root = tree.getroot()
# child = root.find('m:oMath', ns_map)
# print(child.tag, child.attrib)
# text1 = ''
# for idx, equa in enumerate(equations):
#     e = tag_to_latex(equa)
#     print(e, idx)
#     text1 += '&'
#     text1 += tag_to_latex(equa)

#     text1 += '&'


# print(text1)

# xmlstr = ElementTree.tostring(equations[939], encoding='unicode', method='xml')
# print(xmlstr)
# print(type(xmlstr))
# xmlStr = '''<MathEquation>''' + xmlstr + '''</MathEquation>'''  


# for omath in omml.load_string(xmlStr):
#     print(omath.latex)