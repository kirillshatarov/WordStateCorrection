import time
from typing import Any, List
from docx.enum.text import WD_ALIGN_PARAGRAPH
import os
import docx

from file_reader import file_reader

setter = {
    "по ширине": WD_ALIGN_PARAGRAPH.JUSTIFY, "по левому краю": WD_ALIGN_PARAGRAPH.LEFT,
    "по правому краю": WD_ALIGN_PARAGRAPH.RIGHT, "по центру": WD_ALIGN_PARAGRAPH.CENTER,
    "по умолчанию": WD_ALIGN_PARAGRAPH is None
}

setter_gost = {
    "JUSTIFY": WD_ALIGN_PARAGRAPH.JUSTIFY, "LEFT": WD_ALIGN_PARAGRAPH.LEFT,
    "RIGHT": WD_ALIGN_PARAGRAPH.RIGHT, "CENTER": WD_ALIGN_PARAGRAPH.CENTER,
    "None": WD_ALIGN_PARAGRAPH is None
}

TITLE = {'Заголовок': 'Title',
         'Заголовок 1': 'Heading 1',
         'Заголовок 2': 'Heading 2',
         'Заголовок 3': 'Heading 3',
         'Заголовок 4': 'Heading 4'}


chekout = ['Выравнивание','Межстрочный интервал','Абзацный отступ']


class FileManger:
    def __init__(self, user_id: int, docx_: docx.Document, name: str, rej: int = None):
        self.user_id = user_id
        self.user_file = docx_
        self.user_file_name = name
        self.alignment = None
        self.indent = None
        self.interval = None
        self.gost = None
        self.rej = rej

    def checker(self):
        if self.rej == 0:
            self.alignment = setter[self.alignment]
            return self.align()
        elif self.rej == 1:
            return self.catcher()
        elif self.rej == 2:
            return self.checkIndents()
        elif self.rej == -1:
            func = [self.align(),
                    self.catcher(),
                    self.checkIndents()]
            return '\n'.join([f'Проверка на {chekout[i]} \n\n{fnc}\n\n' for i, fnc in enumerate(func)])

    def full_check(self):
        if self.gost is not None:
            params = file_reader(self.gost + '.json').read_file()
            self.alignment = setter_gost[params['alignment']]
            self.indent = params['paragraph-indent']
            self.interval = params['interval']
            self.rej = -1
            return self.checker()
        print('sa')
        return None

    def line_space(self) -> list[float]:
        paragraphs = self.user_file.paragraphs
        return [paragraph.paragraph_format.line_spacing for paragraph in paragraphs]

    def catcher(self) -> str:
        errors = []
        for i, p_ in enumerate(self.line_space()):
            if isinstance(self.interval,list):
                if self.interval[0] <= p_ <= self.interval[1]:
                    errors.append(f"Не соответствие госту в {i + 1} строке")
            else:
                if p_ != self.interval:
                    errors.append(f"Не соответствие госту в {i + 1} строке")
        if errors:
            return "\n".join(errors)
        else:
            return "Все соответствует госту"

    def is_fully_alignment(self) -> list[bool]:
        paragraphs = self.user_file.paragraphs
        return [paragraph.alignment != self.alignment for paragraph in paragraphs
                if paragraph.style.name not in ["Heading 1", "Heading 2"]]

    def align(self) -> str:
        errors = []
        for i, p_ in enumerate(self.is_fully_alignment()):
            if p_:
                errors.append(f"Не соответствие госту в {i + 1} строке")
        if errors:
            return "\n".join(errors)
        else:
            return "Все соответствует госту"

    def checkIndents(self):
        document = self.user_file
        paragraphs = document.paragraph
        result = ""
        for paragraph in paragraphs:
            # Проверяем, не является ли абзац заголовком
            if paragraph.style.name not in TITLE.values():
                # проверяем, является ли абзац - списоком, если да, то идем дальше по документу.
                if len(paragraph._element.xpath('./w:pPr/w:numPr')) > 0:
                    continue
                if paragraph.paragraph_format.first_line_indent is not None:
                    doc_indent = round(paragraph.paragraph_format.first_line_indent.cm, 2)
                    if isinstance(self.indent,list):
                        left = self.indent[0]
                        right = self.indent[1]
                        if not (left <= doc_indent <= right):
                            result += f'Абзац: "{paragraph.text[:25]}" оформлен неверно. Его отступ ' \
                                      f'составляет: {doc_indent} см.\n-----\n'
                    else:
                        if doc_indent != float(self.indent):
                            result += f'Абзац: "{paragraph.text[:25]}" оформлен неверно. Его отступ ' \
                                  f'составляет: {doc_indent} см.\n-----\n'

        if result.count('\n') == 0:
            return "Все отступы оформлены верно."
        return result

    def close(self):
        os.remove(self.user_file_name)
