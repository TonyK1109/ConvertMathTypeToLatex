import json
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
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    '': 'http://schemas.openxmlformats.org/package/2006/relationships',
}


def remove_escape_sequence(string):
    return string.strip('\t').strip('\n').strip('\n\n')


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


def get_image_location_dict(string):
    dict = {}
    relationships = ElementTree.fromstring(string)
    for child in relationships.findall('Relationship', ns_map):
        dict[child.attrib['Id']] = child.attrib['Target']
    return dict


def get_image_location(image_dict, tag):
    for child in tag.iter(qn('a:blip')):
        rid = child.attrib[qn('r:embed')]
        return image_dict[rid]


def get_answer_table_string(tag):
    text = ''
    for child in tag.iter():
        if child.tag == qn('w:t'):
            text += child.text if child.text is not None else ''
        elif child.tag == qn('w:tab'):
            text += '\t'
        elif child.tag == qn('w:br') or child.tag == qn('w:cr'):
            text += '\n'
        elif child.tag == qn('w:p'):
            text += '\n\n'

    return text


def get_answer_dict(text):
    answer_dict = {}

    table_string = text.split('BẢNG ĐÁP ÁN VÀ HƯỚNG DẪN GIẢI')[
        1].split('&t')[1]
    answer_list_row1 = table_string.split('\n\n')[26:51]
    answer_list_row2 = table_string.split('\n\n')[76:101]
    answer_list_row1.extend(answer_list_row2)
    for i in range(50):
        answer_dict[i+1] = answer_list_row1[i]

    return answer_dict


def get_model_question(answer_dict, index, question):

    body = remove_escape_sequence(question.split('A.')[0])
    question_answer_dict = {}

    question_answer_dict['A'] = remove_escape_sequence(
        question.split('A. ')[1].split('B. ')[0])
    question_answer_dict['B'] = remove_escape_sequence(
        question.split('A. ')[1].split('B. ')[1].split('C. ')[0])
    question_answer_dict['C'] = remove_escape_sequence(
        question.split('A. ')[1].split('B. ')[1].split('C. ')[1].split('D. ')[0])
    question_answer_dict['D'] = remove_escape_sequence(
        question.split('A. ')[1].split('B. ')[1].split('C. ')[1].split('D. ')[1])

    answers = []
    for key, value in question_answer_dict.items():
        if key == answer_dict[index]:
            answer = {"value": value, "isCorrect": True}
        else:
            answer = {"value": value}
        answers.append(answer)

    return {
        "subject": "Toán",
        "body": body,
        "answers": answers
    }


def get_model_question_list(text, answer_dict):
    model_question_list = []

    question_list = text.split('----HẾT----')[0].split('&*')
    index_question = 1  # start with 1
    for question in question_list[1:51]:
        model_question_list.append(get_model_question(
            answer_dict, index_question, question))
        index_question += 1

    return model_question_list


# Main Flow
def convert_from_word_to_model_question_list(path_file):
    zip_f = zipfile.ZipFile(path_file)

    for f in zip_f.namelist():
        if f.startswith('word/document'):
            document = zip_f.read(f)
        if f == 'word/_rels/document.xml.rels':
            relations = zip_f.read(f)

    # save image
    for f in zip_f.namelist():
        _, extension = os.path.splitext(f)
        if extension in extensions:
            destination = os.path.join("image", os.path.basename(f))
            with open(destination, 'wb') as destination_file:
                destination_file.write(zip_f.read(f))

    zip_f.close()

    # load relationships for get image location
    image_dict = get_image_location_dict(relations)

    text = ''
    root = ElementTree.fromstring(document)
    for child in root.iter():

        if child.tag == qn('w:t'):
            text += child.text if child.text is not None else ''

        # divide MCQ
        elif child.tag == qn('w:bookmarkStart'):
            text += '&*'

        # Found answer table
        elif child.tag == qn('w:tbl'):
            text += '&t'
            text += get_answer_table_string(child)
            text += '&t'

        # Found an equation
        elif child.tag == qn('m:oMath'):
            text += inline_delimiter + ' '
            text += tag_to_latex(child)
            text += ' ' + inline_delimiter

        # Found an image
        elif child.tag == qn('w:drawing'):
            url = get_image_location(image_dict, child)
            if url is not None:
                text += f'\nIMAGE_URL:{url}\n'

        elif child.tag == qn('w:tab'):
            text += '\t'
        elif child.tag == qn('w:br') or child.tag == qn('w:cr'):
            text += '\n'
        elif child.tag == qn('w:p'):
            text += '\n\n'

    text = re.sub(r'\n(\n+)\$(\s*.+\s*)\$\n', r'\n\1$$ \2 $$', text)

    answer_dict = get_answer_dict(text)

    model_question_list = get_model_question_list(text, answer_dict)

    with open('model_question_list.json', 'w', encoding='utf-8') as f:
        json.dump(model_question_list, f, ensure_ascii=False,  indent=4)


if __name__ == '__main__':
    convert_from_word_to_model_question_list('Original.docx')
