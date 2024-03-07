from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import RGBColor

setter_gost = {
    "по ширине": WD_ALIGN_PARAGRAPH.JUSTIFY, "по левому краю": WD_ALIGN_PARAGRAPH.LEFT,
    "по правому краю": WD_ALIGN_PARAGRAPH.RIGHT, "по центру": WD_ALIGN_PARAGRAPH.CENTER,
    "по умолчанию": WD_ALIGN_PARAGRAPH is None
}

setter_eng = {
    "JUSTIFY": WD_ALIGN_PARAGRAPH.JUSTIFY, "LEFT": WD_ALIGN_PARAGRAPH.LEFT,
    "RIGHT": WD_ALIGN_PARAGRAPH.RIGHT, "CENTER": WD_ALIGN_PARAGRAPH.CENTER,
    "None": WD_ALIGN_PARAGRAPH is None
}

TITLE = {'Заголовок': 'Title',
         'Заголовок 1': 'Heading 1',
         'Заголовок 2': 'Heading 2',
         'Заголовок 3': 'Heading 3',
         'Заголовок 4': 'Heading 4'}

chekout = ['Выравнивание', 'Межстрочный интервал', 'Абзацный отступ']

colors = {
    'red': RGBColor(255, 0, 0),
    'green': RGBColor(0, 255, 0),
    'blue': RGBColor(0, 0, 255),
    'yellow': RGBColor(196, 120, 20),
    'pink': RGBColor(250, 105, 165)
}

exceptions = {
    'line_spacing': 'Межстрочный интервал не соответствует ГОСТу.\nОн должен быть: ',
    'alignment': 'Выравнивание не соответствует ГОСТу.\nОно должно быть: ',
    'indent': 'Абзацный отступ не соответствует ГОСТу.\nОн должен быть: ',
    'font-size': 'Размер шрифта не соответствует ГОСТу.\nОн должен быть: ',
    'font-style': 'Стиль шрифта не соответствует ГОСТу.\nОн должен быть: '
}

color_exceptions = {
    'green': 'alignment',
    'blue': 'line_spacing',
    'yellow': 'indent'
}
