import os

import docx

import docx_ed.cfg as c
from docx_ed.file_reader import file_reader


def painter(paragraph: docx, errors: list[tuple]):
    for color, comment in errors:
        paint(paragraph, color, comment)


def paint(paragraph: docx, color: str, comment: str):
    paragraph.add_comment(comment)
    for word in paragraph.runs:
        word.font.color.rgb = c.colors[color]


class FileManger:
    def __init__(self, user_id: int, docx_: docx.Document, name: str, doc_rej: bool = False, bot_rej: int = None,
                 gost=None):
        self.gost = gost
        self.user_id = user_id
        self.user_file = docx_
        self.user_file_name = name
        self.alignment = None
        self.indent = None
        self.interval = None
        self.doc_rej = doc_rej
        self.bot_rej = bot_rej

    @staticmethod
    def msg_errors(errors: list) -> str:
        if errors:
            return "\n".join(errors)
        return "Все соответствует госту"

    def is_fully_alignment(self) -> str:
        paragraphs = self.user_file.paragraphs
        errors = []
        for i, paragraph in enumerate(paragraphs):
            if paragraph.alignment != self.alignment:
                errors.append(f"Не соответствие госту в {i + 1} строке")
        return self.msg_errors(errors)

    def is_correct_indents(self):
        paragraphs = self.user_file.paragraphs
        errors = []
        for i, paragraph in enumerate(paragraphs):
            if paragraph.style.name not in c.TITLE.values():
                if len(paragraph._element.xpath('./w:pPr/w:numPr')) > 0:
                    continue
                if paragraph.paragraph_format.first_line_indent is not None:
                    doc_indent = round(paragraph.paragraph_format.first_line_indent.cm, 2)
                    if isinstance(self.indent, list):
                        left, right = self.indent
                        if not (left <= doc_indent <= right):
                            errors.append(f'Абзац: "{paragraph.text[:25]}" оформлен неверно. Его отступ '
                                          f'составляет: {doc_indent} см')
                    else:
                        if doc_indent != float(self.indent):
                            errors.append(f'Абзац: "{paragraph.text[:25]}" оформлен неверно. Его отступ '
                                          f'составляет: {doc_indent} см')
        return self.msg_errors(errors)

    def is_correct_line_spacing(self):
        lines = [paragraph.paragraph_format.line_spacing for paragraph in self.user_file.paragraphs if
                 paragraph.style.name not in c.TITLE.values()]
        errors = []
        for i, p_ in enumerate(lines):
            if isinstance(self.interval, list):
                if self.interval[0] <= p_ <= self.interval[1]:
                    errors.append(f"Не соответствие госту в {i + 1} строке")
            else:
                if p_ != self.interval:
                    errors.append(f"Не соответствие госту в {i + 1} строке")
        return self.msg_errors(errors)

    def get_params_from_ghost(self):
        if self.gost in file_reader.get_files().keys():
            params = file_reader(self.gost + '.json').read_file()
            self.alignment = c.setter_gost[params['alignment']]
            self.indent = params['paragraph-indent']
            self.interval = params['interval']
            self.rej = -1
            return True
        return False

    def is_correct_document(self):
        if self.get_params_from_ghost():
            errors = {'alignment': [],
                      'line_spacing': [],
                      'indent': []
                      }
            for i, paragraph in enumerate(self.user_file.paragraphs):
                if not (paragraph.style.name.startswith('Heading') and
                        paragraph.style.name.startswith('Subheading')):
                    # Для списков
                    par_errors = []
                    if len(paragraph._element.xpath('./w:pPr/w:numPr')) > 0:
                        continue
                    # Выравнивание
                    if paragraph.alignment != self.alignment:
                        comment = c.exceptions['alignment'] + str(self.alignment)
                        par_errors.append(('green', comment))

                    # Межстрочный интервал
                    interval = paragraph.paragraph_format.line_spacing
                    if interval:
                        if isinstance(self.interval, list):
                            left, right = map(float, self.indent)
                            if left <= interval <= right:
                                comment = c.exceptions['line_spacing'] + '-'.join(self.interval)
                                par_errors.append(('yellow', comment))

                        else:
                            if interval != self.interval:
                                comment = c.exceptions['line_spacing'] + str(self.interval)
                                par_errors.append(('yellow', comment))

                    # абзацы скипаем
                    if paragraph.paragraph_format.first_line_indent is not None:
                        # Отступ
                        doc_indent = round(paragraph.paragraph_format.first_line_indent.cm, 2)
                        if isinstance(self.indent, list):
                            left, right = map(float, self.indent)
                            if not (left <= doc_indent <= right):
                                comment = c.exceptions['indent'] + '-'.join(self.indent)
                                par_errors.append(('blue', comment))
                        else:
                            if doc_indent != float(self.indent):
                                comment = c.exceptions['indent'] + str(self.indent)
                                par_errors.append(('blue', comment))

                    if par_errors:
                        if self.doc_rej:
                            painter(paragraph, par_errors)
                        else:
                            for color, comment in par_errors:
                                errors[c.color_exceptions[color]].append(comment + f'\n На строке {i}')
            return self.answer(errors)
        return False

    def answer(self, errors: dict = None):
        if self.doc_rej:
            self.saver()
        else:
            answer = ''
            for keyh in errors:
                if errors[keyh]:
                    answer += f'Проблемы возникли с {keyh}: \n' + '\n'.join(errors[keyh]) + '\n'
            if answer:
                print(answer)
                return answer
            else:
                return 'Все соотвествует ГОСТу'

    def saver(self):
        self.user_file.save(f'../files/edited_Docx/{self.user_id}_ready_file.docx')

    def close(self):
        os.remove(self.user_file_name)


obj = FileManger(1, docx.Document('../test2.docx'), 'tur', gost="2.105-2019", doc_rej=True)
obj.is_correct_document()
obj.user_file.save('Исправлен.docx')
