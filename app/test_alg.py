import json
import sys

import docx
from PyQt5 import QtCore, QtGui
from PyQt5.QtGui import QPainter, QColor, QBrush, QRegExpValidator
from PyQt5.QtCore import QSize, QRegExp
from PyQt5.QtWidgets import (QApplication, QLabel, QPushButton, QComboBox, QLineEdit, QHBoxLayout, QVBoxLayout,
                             QGridLayout,
                             QWidget, QScrollArea, QPlainTextEdit, QMessageBox, QFileDialog, QMainWindow, QSizePolicy)
from PyQt5.QtCore import QFile, QTextStream

from constants import READ_ONLY, SETTER
from docx_cls import FileManger
from secondWindow import SecondWindow


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Проверка файла")
        self.setGeometry(200, 40, 1440, 990)
        # self.setMinimumSize(800, 600)
        # self.setMaximumSize(QSize(1940, 990))
        self.setMinimumSize(QSize(820, 820))  # 980, 800
        self.pathFile = ''
        # self.second_window = None
        # self.plain_text = None


        self.central_widget = QWidget()
        # size_policy = QSizePolicy(QSizePolicy.MinimumExpanding, QSizePolicy.MinimumExpanding)
        self.setCentralWidget(self.central_widget)
        # self.central_widget.setSizePolicy(size_policy)

        self.initUI()


    def initUI(self):
        self.title = QPlainTextEdit('Проверка файла по своим настройкам', self)
        self.title.setReadOnly(READ_ONLY)
        # self.title = QLabel('Проверка файла по своим настройкам', self)
        self.title.setMinimumSize(200, 80)
        self.title.setMaximumSize(700, 100)
        # self.title.setGeometry(QtCore.QRect(20, 20, 502, 50))
        self.title.setStyleSheet('''
                        QPlainTextEdit {
                                          border: 1px solid #0074BA;
                                          /* background-color: #0074BA; */
                                          font-size: 28px;
                                          font-weight: 700;
                                          font-family: 'Aleo';
                                          color: #FFFFFF;
                                          /* width: 30%; height: 60%; */
                                          /* padding: 0px 20px 0px 0px; */
                                    }
                                ''')

        # Кнопка для открывания второго окна
        self.window2_button = QPushButton("Проверить по ГОСТ", self)
        self.window2_button.setMinimumSize(200, 50)
        # self.window2_button.setGeometry(QtCore.QRect(836, 0, 604, 69))
        self.window2_button.setStyleSheet('''
                    QPushButton {
                              font-weight: 700;
                              background-color: #E9E9E9;
                              font-family: 'Aleo';
                              font-size: 20px;
                              color: #000000;
                              /* width: 50%; height: 70%; */
                    }
                    ''')


        self.pickAligment = QComboBox(self)
        # self.pickAligment.setGeometry(QtCore.QRect(20, 200, 150, 31))
        self.pickAligment.addItems(SETTER.keys())
        self.pickAligment.setMinimumSize(200, 50)
        self.pickAligment.setMaximumSize(90, 50)
        self.pickAligment.setStyleSheet('''
                        QComboBox {
                                    font-size: 20px;
                                    font-weight: 400;
                                    font-family: 'Aleo';
                                    color: #000000;
                                    background-color: #FFFFFF;
                                    /* margin: 0px 30px 0px 0px; */
                                    border: 2px solid transparent; /* Прозрачная рамка */
                                    border-radius: 6px;
                                    padding: 0px 0px 0px 10px; /* Внутренний отступ */
                                    /* width: 30%; height: 60%; */
                                }
           QComboBox::drop-down {
                                    subcontrol-position: right;
                                    /* width: 25%; */
                                    /*padding: 1px;*/
                                }
    QComboBox QAbstractItemView {
                                    font-size: 20px;
                                    font-weight: 400;
                                    font-family: 'Aleo';
                                    color: #000000;
                                    background-color: #FFFFFF;
                                    /* margin: 0px 30px 0px 0px; */
                                    selection-background-color: #C0C0C0; /* Цвет фона при наведении на элемент списка */
                                    selection-color: #000000; /* Цвет текста при наведении на элемент списка */
                                    /*padding: 30px 5px;*/
                                }
    QComboBox QAbstractItemView::item {
                                    padding: 30px 30px;
                                }
                        ''')
        # self.pickAligment.view().setStyleSheet('''
        #    QComboBox QAbstractItemView::item {
        #         padding: 5px 10px; /* Задайте здесь нужные вам отступы */
        #     }
        # ''')


        # self.labelAlignment = QLabel('Выравнивание', self)
        self.labelAlignment = QPlainTextEdit('Выравнивание абзаца', self)
        self.labelAlignment.setReadOnly(READ_ONLY)
        # self.labelAlignment.setGeometry(QtCore.QRect(20, 160, 250, 31))
        # self.labelAlignment.setStyleSheet("color: #FFFFFF;")
        self.labelAlignment.setMaximumSize(220, 40)
        self.labelAlignment.setMinimumSize(100, 61)
        self.labelAlignment.setStyleSheet('''
                                    QPlainTextEdit {
                                            border: 2px solid red;
                                            font-size: 20px;
                                            font-family: 'Aleo';
                                            color: #F4F2F2;
                                            /* margin: 30px 30px 0px 0px; */
                                            /* padding: 30px 0 10px 1px; */
                                    }
                                    ''')

        # self.pickIndent = QLabel('Отступ', self)
        self.pickIndent = QPlainTextEdit('Отступ', self)
        self.pickIndent.setReadOnly(READ_ONLY)
        # self.pickIndent.setGeometry(QtCore.QRect(350, 250, 270, 31))
        # self.pickIndent.setStyleSheet("color: #FFFFFF;")
        self.pickIndent.setMaximumSize(700, 40)
        self.pickIndent.setMinimumSize(100, 40)
        self.pickIndent.setStyleSheet('''
                            QPlainTextEdit {
                                    border: 2px solid red;
                                    font-size: 20px;
                                    font-family: 'Aleo';
                                    color: #F4F2F2;
                                    /* margin: 30px 30px 0px 0px; */
                                    /* padding: 30px 0 10px 1px; */
                            }
                            ''')

        self.enterIndent = QLineEdit(self)
        # self.enterIndent.setGeometry(QtCore.QRect(350, 290, 80, 31))
        self.enterIndent.setPlaceholderText('0 см')
        self.enterIndent.setValidator(QtGui.QDoubleValidator())
        # self.enterIndent.setMinimumSize(200, 50)
        self.enterIndent.setMaximumSize(100, 50)
        self.enterIndent.setStyleSheet('''
                                        QLineEdit {
                                                align-text: center;
                                                font-size: 20px;
                                                font-weight: 400;
                                                background-color: #FFFFFF;
                                                font-family: 'Aleo';
                                                color: #000000;
                                                /* margin: 0px 30px 0px 0px; */
                                                padding: 0px 0px 0px 10px;
                                                border-radius: 6px;
                                                /* width: 30%; height: 60%; */
                                        }
                                    ''')

        self.enterLineSpace = QLineEdit(self)
        # self.enterLineSpace.setGeometry(QtCore.QRect(350, 200, 80, 31))
        self.enterLineSpace.setPlaceholderText('1 см')
        self.enterLineSpace.setValidator(QtGui.QDoubleValidator())
        # self.enterLineSpace.setMinimumSize(200, 50)
        self.enterLineSpace.setMaximumSize(90, 50)
        self.enterLineSpace.setStyleSheet('''
                                        QLineEdit {
                                                font-size: 20px;
                                                font-weight: 400;
                                                background-color: #FFFFFF;
                                                font-family: 'Aleo';
                                                color: #000000;
                                                /* margin: 0px 30px 0px 0px; */
                                                padding: 0px 0px 0px 10px;
                                                border-radius: 6px;
                                                /* width: 30%; height: 60%; */
                                        }
                                    ''')

        # self.pickLineSpace = QLabel('Межстрочный интервал', self)
        self.pickLineSpace = QPlainTextEdit('Межстрочный интервал', self)
        self.pickLineSpace.setReadOnly(READ_ONLY)
        # self.pickLineSpace.setGeometry(QtCore.QRect(350, 160, 410, 40))
        # self.pickLineSpace.setStyleSheet("color: #FFFFFF;")
        self.pickLineSpace.setMaximumSize(700, 40)
        self.pickLineSpace.setMinimumSize(100, 61)
        self.pickLineSpace.setStyleSheet('''
                                            QPlainTextEdit {
                                                    border: 2px solid red;
                                                    font-size: 20px;
                                                    font-family: 'Aleo';
                                                    color: #F4F2F2;
                                                    /* margin: 30px 30px 0px 0px; */
                                                    /* padding: 30px 0 10px 1px; */
                                            }
                                            ''')


        self.enterFont = QLineEdit(self)
        # self.enterFont.setGeometry(QtCore.QRect(20, 110, 240, 35))
        # self.enterFont.setStyleSheet("color: #FFFFFF;")
        validator = QRegExpValidator(QRegExp("[A-Za-z]+"))
        self.enterFont.setValidator(validator)
        # self.enterFont.setMinimumSize(200, 50)
        self.enterFont.setMaximumSize(200, 50)
        self.enterFont.setStyleSheet('''
                                        QLineEdit {
                                                font-size: 20px;
                                                font-weight: 400;
                                                background-color: #FFFFFF;
                                                font-family: 'Aleo';
                                                color: #000000;
                                                /* margin: 0px 30px 0px 0px; */
                                                padding: 0px 0px 0px 10px;
                                                border-radius: 6px;
                                                /* width: 30%; height: 60%; */
                                        }
                                    ''')

        # self.pickFont = QLabel('Стиль шрифта', self)
        self.pickFont = QPlainTextEdit('Стиль шрифта', self)
        self.pickFont.setReadOnly(READ_ONLY)
        # self.pickFont.setGeometry(QtCore.QRect(20, 70, 270, 31))
        # self.pickFont.setStyleSheet("color: #FFFFFF;")
        self.pickFont.setMaximumSize(200, 40)
        self.pickFont.setMinimumSize(100, 40)
        self.pickFont.setStyleSheet('''
                                QPlainTextEdit {
                                                    border: 2px solid red;
                                                    font-size: 20px;
                                                    font-family: 'Aleo';
                                                    color: #F4F2F2;
                                                    /* margin: 30px 30px 0px 0px; */
                                                    /* padding: 30px 0 10px 1px; */
                                                }
                                            ''')

        self.enterFontSize = QLineEdit(self)
        # self.enterFontSize.setGeometry(QtCore.QRect(350, 110, 240, 35))
        self.enterFontSize.setPlaceholderText('14')
        # self.enterFontSize.setMinimumSize(80, 50)
        self.enterFontSize.setMaximumSize(90, 50)
        self.enterFontSize.setValidator(QtGui.QIntValidator())
        self.enterFontSize.setStyleSheet('''
                        QLineEdit {
                                        font-size: 20px;
                                        font-weight: 400;
                                        background-color: #FFFFFF;
                                        font-family: 'Aleo';
                                        color: #000000;
                                        /* margin: 0px 30px 0px 0px; */
                                        padding: 0px 0px 0px 10px;
                                        border-radius: 6px;
                                        /* width: 30%; height: 60%; */
                                }
                            ''')

        # self.pickFontSize = QLabel('Размер шрифта', self)
        self.pickFontSize = QPlainTextEdit('Размер шрифта', self)
        self.pickFontSize.setReadOnly(READ_ONLY)
        self.pickFontSize.setMinimumSize(100, 40)
        self.pickFontSize.setMaximumSize(700, 40)
        # self.pickFontSize.setGeometry(QtCore.QRect(350, 70, 270, 31))
        # self.pickFontSize.setStyleSheet("color: #FFFFFF;")
        self.pickFontSize.setStyleSheet('''
                                QPlainTextEdit {
                                                    border: 2px solid red;
                                                    font-size: 20px;
                                                    font-family: 'Aleo';
                                                    color: #F4F2F2;
                                                    /* margin: 30px 30px 0px 0px; */
                                                    /* padding: 30px 0 10px 1px; */
                                                }
                                            ''')

        self.filePicked = QLabel("", self)
        # self.filePicked = QPlainTextEdit('', self)
        # self.filePicked.setReadOnly(READ_ONLY)
        self.filePicked.setMinimumSize(150, 20)
        # self.filePicked.setGeometry(QtCore.QRect(226, 807, 400, 31))
        self.filePicked.setMinimumSize(250, 40)
        self.filePicked.setMaximumSize(420, 40)
        self.filePicked.setStyleSheet('''
                                QLabel {
                                                    border: 2px solid red;
                                                    font-size: 20px;
                                                    font-family: 'Aleo';
                                                    color: #F4F2F2;
                                                    /* border: 1px solid red */
                                                }
                                            ''')
        self.filePicked.setAlignment(QtCore.Qt.AlignCenter)


        # self.pickFile = QLabel('Выберите файл (docx):', self)
        # # self.pickFile.setGeometry(QtCore.QRect(226, 656, 250, 31))
        # # self.pickFile.setStyleSheet("color: #FFFFFF;")
        # self.pickFile.setStyleSheet('''
        #                                 QLabel {
        #                                             font-size: 20px;
        #                                             font-weight: 400;
        #                                             font-family: 'Aleo';
        #                                             color: #FFFFFF;
        #                                         }
        #                                         ''')

        # Кнопка выбора файла
        self.pickFileButton = QPushButton("Выбрать файл (docx)", self)
        # self.pickFileButton.setGeometry(QtCore.QRect(226, 723, 250, 71))
        self.pickFileButton.setMinimumSize(100, 72)
        self.pickFileButton.setMaximumSize(400, 70)
        self.pickFileButton.setAcceptDrops(True)
        self.pickFileButton.setStyleSheet('''
                            QPushButton {
                                          font-weight: 400;
                                          font-family: 'Aleo';
                                          font-size: 20px;
                                          color: #FFFFFF;
                                          text-align: center;
                                          border: 3px solid #FFFFFF;
                                          border-radius: 36px;
                                          padding: 10px 30px;
                                          /* margin: 30px 170px 0px 170px; */
                                          /* width: 30%; height: 50%; */
                                    }
                                ''')

        self.checkFile = QPushButton('Проверить файл', self)
        # self.checkFile.setGeometry(QtCore.QRect(226, 883, 250, 71))
        self.checkFile.setMinimumSize(100, 72)
        self.checkFile.setMaximumSize(400, 70)
        self.checkFile.setStyleSheet('''
                            QPushButton {
                                      font-weight: 400;
                                      font-family: 'Aleo';
                                      font-size: 20px;
                                      color: #FFFFFF;
                                      text-align: center;
                                      border: 3px solid #FFFFFF;
                                      border-radius: 36px;
                                      padding: 10px 30px;
                                      /* margin-bottom: 30px; */
                                      /* margin: 30px 170px 25px 170px; */
                                      /* width: 30%; height: 50%; */
                            }
                            ''')


        self.confirm_button = QPushButton('Подтвердить настройки', self)
        # self.confirm_button.setGeometry(QtCore.QRect(20, 270, 310, 50))
        self.confirm_button.setMinimumSize(100, 72)
        self.confirm_button.setMaximumSize(400, 70)
        self.confirm_button.setStyleSheet('''
                            QPushButton {
                                      font-weight: 400;
                                      font-family: 'Aleo';
                                      font-size: 20px;
                                      color: #FFFFFF;
                                      text-align: center;
                                      border: 3px solid #FFFFFF;
                                      border-radius: 36px;
                                      padding: 10px 30px;
                                      /* margin: 30px 170px 0px 170px; */
                                      /* width: 30%; height: 50%; */
                            }
                            ''')
        self.filename_settings = "My settings"  # название файла со своими настройками проверки



########### ВАЖНО ################

        # self.answer = QScrollArea(self)
        # # self.answer.setGeometry(QtCore.QRect(836, 69, 604, 851))
        # self.answer.setMinimumSize(400, 80)
        # self.answer.setWidgetResizable(True)
        # # self.answer.setAlignment(QtCore.Qt.AlignLeading | QtCore.Qt.AlignLeft | QtCore.Qt.AlignTop)
        # self.answer.setStyleSheet('''
        #                             QScrollArea {
        #                                     padding: 20px;
        #                                     background-color: #FFFFFF;
        #                                     /* min-width: 500px; */
        #                                     /* margin-right: 0px; */
        #                                     /* width: 500%; */
        #                             }
        #                         ''')

########### ВАЖНО ################


        # self.scroll_widget = QWidget()
        # self.scroll_layout = QVBoxLayout(self.scroll_widget)
        #
        # self.plain_text = QPlainTextEdit()
        # self.plain_text.setReadOnly(READ_ONLY)
        # self.scroll_layout.addWidget(self.plain_text)
        #
        # self.scroll_widget.setLayout(self.scroll_layout)
        # self.answer.setWidget(self.scroll_widget)


        ####### МБ НАДО ###########

        self.plain_text = QPlainTextEdit()
        self.plain_text.setReadOnly(READ_ONLY)
        # self.plain_text.setGeometry(QtCore.QRect(836, 69, 604, 851))
        self.plain_text.setMinimumSize(300, 80)
        self.plain_text.setStyleSheet('''
                            QPlainTextEdit {
                                    background-color: #FFFFFF;
                                    border: 4px solid black;
                                    border-radius: 50px;
                                    padding: 25px 10px 25px 40px;
                                    color: #000000;
                                    font-size: 20px;
                                    font-weight: 400;
                                    /* min-width: 200px; */
                                    /* margin: 20px 40px; */
                                    /* width: 500%; */
                            }
                        ''')

        self.downloadFile = QPushButton("Скачать проверенный\nфайл", self)
        # self.downloadFile.setGeometry(836, 920, 604, 69)
        self.downloadFile.setMinimumSize(300, 80)
        self.downloadFile.setMaximumSize(400, 80)
        self.downloadFile.setStyleSheet('''
                                    QPushButton {
                                            font-weight: 700;
                                            background-color: #E9E9E9;
                                            font-family: 'Aleo';
                                            font-size: 20px;
                                            color: #000000;
                                            border: 3px solid #FFFFFF;
                                            border-radius: 36px;
                                            /* margin: 30px 50px 0px 50px; */
                                            padding: 10px 30px;
                                            /* width: 70%; height: 50%; */
                                        }
                                    ''')

        # self.scrollAreaWidgetContents = QWidget()
        # self.scrollAreaWidgetContents.setGeometry(QtCore.QRect(0, 0, 230, 119))


        ########### МБ НАДО ##############

        # layout_2 = QVBoxLayout(self)
        # layout_2.addWidget(self.plain_text)
        #
        # w = QWidget()                 # возможно не нужно
        # w.setLayout(layout_2)         # возможно не нужно
        # self.answer.setWidget(w)      # возможно не нужно
        # w.setStyleSheet("background-color: #FFFFFF;")

        ## self.answer.setLayout(layout_2)

        # ДОБАВЛЕНИЕ ЭЛЕМЕНТОВ В ГРИД #
        grid = QGridLayout(self.central_widget)
        grid.setSpacing(20)
        self.setLayout(grid)
        # grid.setColumnMinimumWidth(0, 50)
        # grid.setColumnMinimumWidth(1, 50)
        # grid.setColumnMinimumWidth(2, 400)

        grid.setContentsMargins(35, 20, 35, 33)
        grid.setColumnStretch(0, 1)  # Установить вес (stretch) для первого столбца
        grid.setColumnStretch(1, 1)  # Установить вес для второго столбца
        grid.setColumnStretch(2, 2)  # Установить вес для третьего столбца
        self.setStyleSheet("background-color: #0074BA;")

        grid.addWidget(self.title, 0, 0, 1, 2)  # title titleQLabel('Проверка файла по своим настройкам')
        grid.addWidget(self.window2_button, 0, 2)  # QPushButton("Проверить по ГОСТ")
        grid.addWidget(self.pickFont, 2, 0)  # QLabel('Стиль шрифта')
        grid.addWidget(self.pickFontSize, 2, 1)  # QLabel('Размер шрифта')
        grid.addWidget(self.enterFont, 3, 0)  # QLineEdit(self) ввести стиль шрифта
        grid.addWidget(self.enterFontSize, 3, 1)  # QLineEdit(self) ввести размер шрифта
        grid.addWidget(self.labelAlignment, 4, 0)  # QLabel('Выравнивание:')
        grid.addWidget(self.pickLineSpace, 4, 1)  # QLabel('Межстрочный интервал')
        grid.addWidget(self.pickAligment, 5, 0)  # QComboBox(self) выравнивание
        grid.addWidget(self.enterLineSpace, 5, 1)  # QLineEdit(self) ввести межстрочный интервал
        grid.addWidget(self.pickIndent, 6, 0)  # QLabel('Отступ')
        grid.addWidget(self.enterIndent, 7, 0)  # QLineEdit(self) ввести отсуп
        grid.addWidget(self.confirm_button, 8, 0, 1, 2, alignment=QtCore.Qt.AlignCenter)  # QPushButton('Подтвердить настройки')
        # grid.addWidget(self.pickFile, 9, 0)  # QLabel('Выберите файл (docx):')
        grid.addWidget(self.pickFileButton, 9, 0, 1, 2, alignment=QtCore.Qt.AlignCenter)  # QPushButton("Выбрать файл")
        grid.addWidget(self.filePicked, 10, 0, 1, 2, alignment=QtCore.Qt.AlignCenter)  # QLabel("", self) выбранный файл
        grid.addWidget(self.checkFile, 11, 0, 1, 2, alignment=QtCore.Qt.AlignCenter)  # QPushButton('Проверить файл')
        # grid.addWidget(self.answer, 1, 2, 11, 1)  # QScrollArea(self)
        grid.addWidget(self.plain_text, 1, 2, 9, 1)  # QPlainTextEdit() поле с ответом
        grid.addWidget(self.downloadFile, 10, 2, 2, 1, alignment=QtCore.Qt.AlignCenter)  # QPushButton("Скачать проверенный файл")
        # grid.addWidget(self.enterIndentLabel, /, / )  # QLabel(' см')
        # grid.addWidget(self.enterLineSpaceLabel, /, / )  # QLabel(' см')


        #
        # события кнопок
        #
        self.pickFileButton.clicked.connect(self.pickFileButton_Clicked)
        self.checkFile.clicked.connect(self.checkFile_Clicked)
        self.window2_button.clicked.connect(self.open_second_window)  # Открытие второго окна для проверки по гостам
        self.confirm_button.clicked.connect(self.confirm_settings)  # Подтверждение настроек гостов
        self.downloadFile.clicked.connect(self.save_ready_file)  # скачивание проверенного файла

        # Подключаем события перетаскивания
        self.pickFileButton.dragEnterEvent = self.dragEnterEvent
        self.pickFileButton.dropEvent = self.dropEvent
        # Добавляем обработку событий перетаскивания файлов
        self.filePicked.setAcceptDrops(True)

        # self.choiceTitle.addItems(TITLE.keys())
        # self.titlePicked.setText(self.choiceTitle.currentText())


    def open_second_window(self):
        self.second_window = SecondWindow(self)
        self.second_window.show()
        self.close()

    def confirm_settings(self):
        # currentIndent = currentIndent.replace(',', '.')
        data = {
            "name": self.filename_settings,
            "font-style": self.enterFont.text(),
            "font-size": self.enterFontSize.text(),
            "paragraph-indent": self.enterIndent.text().replace(',', '.'),
            "interval": self.enterLineSpace.text().replace(',', '.'),
            "alignment": self.pickAligment.currentText()
        }
        # print(self.filename, self.enterFont.text(), self.enterFontSize.text(), self.enterIndent.text(), self.enterLineSpace.text(), self.pickAligment.currentText())
        with open('./files/gost/' + self.filename_settings + '.json', "w", encoding='utf-8') as json_file:
            json.dump(data, json_file, indent=4, ensure_ascii=False)

    def changeIndentLabel(self, text):
        self.enterIndentLabel.setText(text + ' см')

    def changeLineSpaceLabel(self, text):
        self.enterLineSpaceLabel.setText(text + ' см')

    # def choiceTitleActive(self, index):
    #     self.titlePicked.setText(self.choiceTitle.itemText(index))
    #     self.currentTitle = self.titlePicked

    def choiceAlignActive(self, index):
        self.pickAlignmentLabel.setText(self.pickAligment.itemText(index))
        self.currentAlign = self.pickAlignmentLabel

    def pickFileButton_Clicked(self):
        filename, filetype = QFileDialog.getOpenFileName(self,
                                                         "Выбрать файл",
                                                         '.',
                                                         'Word files (*.docx)')
        if filename == '':
            self.filePicked.setText('Файл не выбран.')
            self.pathFile = ''
        else:
            self.pathFile = filename
            filename = filename.split('/')[-1]
            self.filePicked.setText(filename)

    def checkFile_Clicked(self):
        print(self.pathFile)
        if self.pathFile == '':
            try:
                self.notSelectFile = QMessageBox()
                self.notSelectFile.setStandardButtons(QMessageBox.Ok | QMessageBox.Cancel)
                self.notSelectFile.setText('Вы не выбрали файл!')
                self.notSelectFile.setWindowTitle('Ошибка!')
                self.notSelectFile.setIcon(QMessageBox.Warning)
                res = self.notSelectFile.exec()
            except Exception as e:
                pass
        else:
            self.fileName = self.pathFile.split('/')[-1]
            self.path = f'./{self.fileName}'
            print(self.path)
            obj = FileManger(docx.Document(self.path), gost="My settings", doc_rej=False)
            print(1)
            errors = obj.is_correct_document()
            print(2)
            self.plain_text.clear()
            self.plain_text.setPlainText(errors)

    def save_ready_file(self):
        obj2 = FileManger(docx.Document(self.path), gost="My settings", doc_rej=True)
        obj2.is_correct_document()

    def dragEnterEvent(self, event):
        mime_data = event.mimeData()
        if mime_data.hasUrls() and mime_data.urls()[0].isLocalFile():
            event.acceptProposedAction()

    def dropEvent(self, event):
        mime_data = event.mimeData()
        if mime_data.hasUrls() and mime_data.urls()[0].isLocalFile():
            file_path = mime_data.urls()[0].toLocalFile()
            filename = file_path.split('/')[-1]
            # filename = file_path
            self.filePicked.setText(filename)
            self.pathFile = file_path
            event.acceptProposedAction()


# if __name__ == '__main__':
#     app = QApplication(sys.argv)
#     # app.setStyleSheet(stream.readAll())
#     main_window = MainWindow()
#     main_window.show()
#     # main_window.confirm_settings()  # Вызываем сохранение файла JSON перед завершением приложения
#     sys.exit(app.exec_())

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
