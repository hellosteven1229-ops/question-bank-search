import hashlib
import json
import os
import re
import tempfile
import unicodedata
from pathlib import Path


def cell_text(value):
    return '' if value is None else str(value).strip()


def load_excel(path):
    from openpyxl import load_workbook

    workbook = load_workbook(path, data_only=True, read_only=True, keep_links=False)
    questions = []
    try:
        for sheet in workbook.worksheets:
            sheet.reset_dimensions()
            rows = iter(sheet.values)
            headers = {cell_text(value): i for i, value in enumerate(next(rows, ())) if cell_text(value)}
            title_header = next((name for name in ('题目', '题干') if name in headers), None)
            if title_header is None:
                continue
            if not headers.keys() & {'答案', '答案字母', '答案文字'}:
                raise ValueError(f'「{sheet.title}」缺少答案列。请添加“答案”、“答案字母”或“答案文字”。')
            for row_number, row in enumerate(rows, 2):
                def get(name):
                    index = headers.get(name)
                    return cell_text(row[index]) if index is not None and index < len(row) else ''

                title = get(title_header)
                if not title:
                    continue
                options = [{'letter': letter, 'text': value} for letter in 'ABCDEF'
                           if (value := get('选项' + letter) or get(letter))]
                if not options and get('选项'):
                    values = [part.strip() for part in get('选项').split('|')]
                    keys = values if values == ['Y', 'N'] else 'ABCDEF'
                    if len(values) > len(keys):
                        raise ValueError(f'「{sheet.title}」第 {row_number} 行选项超过 6 个。')
                    options = [{'letter': letter, 'text': {'Y': '正确', 'N': '错误'}.get(letter, value)}
                               for letter, value in zip(keys, values) if value]
                raw = get('答案字母') or get('答案')
                answer_text = get('答案文字')
                if not raw and not answer_text:
                    raise ValueError(f'「{sheet.title}」第 {row_number} 行缺少答案。')
                letters, answers = resolve_answer(options, raw, answer_text)
                questions.append(dict(title=title, options=options, letters=letters, answer=answers,
                                      category=get('知识点') or get('类别'), kind=get('题型')))
    finally:
        workbook.close()
    if not questions:
        raise ValueError('未找到题目。请在首行使用“题目”或“题干”，并填写答案列。')
    return questions


def resolve_answer(options, raw, answer_text):
    normalized = unicodedata.normalize('NFKC', raw).upper()
    keys = re.sub(r'[\s,、;|/]+', '', normalized)
    by_letter = {option['letter']: option['text'] for option in options}
    if keys and all(key in by_letter for key in keys):
        correct = set(keys)
        answers = [by_letter[key] for key in dict.fromkeys(keys)]
        letters = '、'.join(dict.fromkeys(keys))
    else:
        value = answer_text or raw
        correct = {letter for letter, text in by_letter.items() if text == value}
        answers = [value]
        letters = '、'.join(sorted(correct))
    for option in options:
        option['correct'] = option['letter'] in correct
    return letters, answers


def default_library():
    return Path(os.getenv('LOCALAPPDATA') or Path.home()) / '题库速查' / '题库'


def write_template(path):
    from openpyxl import Workbook

    workbook = Workbook(write_only=True)
    sheet = workbook.create_sheet('题库')
    sheet.append(['题型', '题目', *(f'选项{letter}' for letter in 'ABCDEF'), '答案', '知识点'])
    workbook.save(path)


def save_bank(directory, path, questions):
    directory.mkdir(parents=True, exist_ok=True)
    identity = os.path.normcase(str(path.resolve()))
    key = hashlib.sha256(identity.encode('utf-8')).hexdigest()
    destination = directory / (key + '.json')
    temp_path = None
    try:
        with tempfile.NamedTemporaryFile('w', encoding='utf-8', dir=directory, delete=False) as out:
            temp_path = Path(out.name)
            json.dump(dict(name=path.stem, questions=questions), out, ensure_ascii=False)
        temp_path.replace(destination)
    finally:
        if temp_path is not None:
            temp_path.unlink(missing_ok=True)
    return destination


def read_banks(directory):
    banks = []
    for path in sorted(directory.glob('*.json')):
        try:
            bank = json.loads(path.read_text(encoding='utf-8'))
            banks.append((path, bank['name'], bank['questions']))
        except (OSError, ValueError, KeyError, TypeError):
            continue
    return banks
