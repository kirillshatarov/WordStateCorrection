import json
import sys

import docx
from PyQt5 import QtCore, QtGui
from PyQt5.QtCore import QSize, QRegExp
from PyQt5.QtGui import QRegExpValidator
from PyQt5.QtWidgets import (QApplication, QLabel, QPushButton, QComboBox, QLineEdit, QGridLayout,
                             QWidget, QScrollArea, QPlainTextEdit, QMessageBox, QFileDialog)

from constants import READ_ONLY, SETTER
from docx_cls import FileManger


class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Проверка файла")
        self.setGeometry(200, 40, 1440, 1024)
        self.setMaximumSize(QSize(1940, 990))
        self.setMinimumSize(QSize(980, 800))
        self.pathFile = ''
        self.second_window = None
        # self.plain_text = None

        self.initUI()


    def initUI(self):
        self.title = QLabel('Проверка файла по своим настройкам', self)
        self.title.setStyleSheet('''
                                  QLabel {
                                          font-size: 29px;
                                          font-weight: 700;
                                          font-family: 'Aleo';
                                          color: #FFFFFF;
                                          width: 30%; height: 60%;
                                          padding: 0px 20px 0px 0px;
                                        }
                                    ''')

        # Кнопка для открывания второго окна
        self.window2_button = QPushButton("Проверить по ГОСТ", self)
        self.window2_button.setStyleSheet('''
                    QPushButton {
                              font-weight: 700;
                              background-color: #E9E9E9;
                              font-family: 'Aleo';
                              font-size: 20px;
                              color: #000000;
                              width: 50%; height: 70%;
                    }
                    ''')


        self.pickAligment = QComboBox(self)
        self.pickAligment.addItems(SETTER.keys())
        self.pickAligment.setStyleSheet('''
                        QComboBox {
                                    font-size: 20px;
                                    font-weight: 400;
                                    font-family: 'Aleo';
                                    color: #000000;
                                    background-color: #FFFFFF;
                                    margin: 0px 30px 0px 0px;
                                    border: 2px solid transparent; /* Прозрачная рамка */
                                    border-radius: 6px;
                                    padding: 0px 0px 0px 10px; /* Внутренний отступ */
                                    width: 30%; height: 60%;
                                }
           QComboBox::drop-down {
                                    subcontrol-position: right;
                                    width: 25%;
                                    /*padding: 1px;*/
                                }
    QComboBox QAbstractItemView {
                                    font-size: 20px;
                                    font-weight: 400;
                                    font-family: 'Aleo';
                                    color: #000000;
                                    background-color: #FFFFFF;
                                    margin: 0px 30px 0px 0px;
                                    selection-background-color: #C0C0C0; /* Цвет фона при наведении на элемент списка */
                                    selection-color: #000000; /* Цвет текста при наведении на элемент списка */
                                    /*padding: 30px 5px;*/
                                }
    QComboBox QAbstractItemView::item {
                                    padding: 30px 30px;
                                }
                        ''')

        self.labelAlignment = QLabel('Выравнивание:', self)
        self.labelAlignment.setStyleSheet('''
                                    QLabel {
                                            font-size: 20px;
                                            font-family: 'Aleo';
                                            color: #F4F2F2;
                                            margin: 30px 30px 0px 0px;
                                            padding: 30px 0 10px 1px;
                                    }
                                    ''')

        self.pickIndent = QLabel('Отступ', self)
        self.pickIndent.setStyleSheet('''
                            QLabel {
                                    font-size: 20px;
                                    font-family: 'Aleo';
                                    color: #F4F2F2;
                                    margin: 30px 30px 0px 0px;
                                    padding: 30px 0 10px 1px;
                            }
                            ''')

        self.enterIndent = QLineEdit(self)
        self.enterIndent.setPlaceholderText('0 см')
        self.enterIndent.setValidator(QtGui.QDoubleValidator())
        self.enterIndent.setStyleSheet('''
                                        QLineEdit {
                                                font-size: 20px;
                                                font-weight: 400;
                                                background-color: #FFFFFF;
                                                font-family: 'Aleo';
                                                color: #000000;
                                                margin: 0px 30px 0px 0px;
                                                padding: 0px 0px 0px 10px;
                                                border-radius: 6px;
                                                width: 30%; height: 60%;
                                        }
                                    ''')

        self.enterLineSpace = QLineEdit(self)
        self.enterLineSpace.setPlaceholderText('1 см')
        self.enterLineSpace.setValidator(QtGui.QDoubleValidator())
        self.enterLineSpace.setStyleSheet('''
                                        QLineEdit {
                                                font-size: 20px;
                                                font-weight: 400;
                                                background-color: #FFFFFF;
                                                font-family: 'Aleo';
                                                color: #000000;
                                                margin: 0px 30px 0px 0px;
                                                padding: 0px 0px 0px 10px;
                                                border-radius: 6px;
                                                width: 30%; height: 60%;
                                        }
                                    ''')

        self.pickLineSpace = QLabel('Межстрочный интервал', self)
        self.pickLineSpace.setStyleSheet('''
                                            QLabel {
                                                    font-size: 20px;
                                                    font-family: 'Aleo';
                                                    color: #F4F2F2;
                                                    margin: 30px 30px 0px 0px;
                                                    padding: 30px 0 10px 1px;
                                            }
                                            ''')


        self.enterFont = QLineEdit(self)
        validator = QRegExpValidator(QRegExp("[A-Za-z]+"))
        self.enterFont.setValidator(validator)
        self.enterFont.setStyleSheet('''
                                        QLineEdit {
                                                font-size: 20px;
                                                font-weight: 400;
                                                background-color: #FFFFFF;
                                                font-family: 'Aleo';
                                                color: #000000;
                                                margin: 0px 30px 0px 0px;
                                                padding: 0px 0px 0px 10px;
                                                border-radius: 6px;
                                                width: 30%; height: 60%;
                                        }
                                    ''')

        self.pickFont = QLabel('Стиль шрифта', self)
        self.pickFont.setStyleSheet('''
                                        QLabel {
                                                    font-size: 20px;
                                                    font-family: 'Aleo';
                                                    color: #F4F2F2;
                                                    margin: 30px 30px 0px 0px;
                                                    padding: 30px 0 10px 1px;
                                                }
                                                ''')

        self.enterFontSize = QLineEdit(self)
        self.enterFontSize.setPlaceholderText('14')
        self.enterFontSize.setValidator(QtGui.QIntValidator())
        self.enterFontSize.setStyleSheet('''
                                QLineEdit {
                                        font-size: 20px;
                                        font-weight: 400;
                                        background-color: #FFFFFF;
                                        font-family: 'Aleo';
                                        color: #000000;
                                        margin: 0px 30px 0px 0px;
                                        padding: 0px 0px 0px 10px;
                                        border-radius: 6px;
                                        width: 30%; height: 60%;
                                }
                            ''')

        self.pickFontSize = QLabel('Размер шрифта', self)
        self.pickFontSize.setStyleSheet('''
                                        QLabel {
                                                    font-size: 20px;
                                                    font-family: 'Aleo';
                                                    color: #F4F2F2;
                                                    margin: 30px 30px 0px 0px;
                                                    padding: 30px 0 10px 1px;
                                                }
                                                ''')

        self.filePicked = QLabel("", self)
        self.filePicked.setStyleSheet('''
                                        QLabel {
                                                    font-size: 20px;
                                                    font-family: 'Aleo';
                                                    color: #F4F2F2;
                                                    /* border: 1px solid red */
                                                }
                                                ''')
        self.filePicked.setAlignment(QtCore.Qt.AlignCenter)

        # Кнопка выбора файла
        self.pickFileButton = QPushButton("Выбрать файл (docx)", self)
        self.pickFileButton.setMinimumSize(600, 80)
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
                                          margin: 30px 170px 0px 170px;
                                          width: 30%; height: 50%;
                                    }
                                ''')

        self.checkFile = QPushButton('Проверить файл', self)
        self.checkFile.setMinimumSize(600, 105)
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
                                      margin: 30px 170px 25px 170px;
                                      width: 30%; height: 50%;
                            }
                            ''')


        self.confirm_button = QPushButton('Подтвердить настройки', self)
        self.confirm_button.setMinimumSize(600, 80)
        self.confirm_button.setStyleSheet('''
                            QPushButton {
                                      font-weight: 400;
                                      font-family: 'Aleo';
                                      font-size: 20px;
                                      color: #FFFFFF;
                                      text-align: center;
                                      border: 3px solid #FFFFFF;
                                      border-radius: 36px;
                                      padding: 10px 10px;
                                      margin: 30px 170px 0px 170px;
                                      width: 30%; height: 50%;
                            }
                            ''')
        self.filename_settings = "My settings"  # название файла со своими настройками проверки


        self.answer = QScrollArea(self)
        self.answer.setMinimumSize(950, 80)
        self.answer.setWidgetResizable(True)
        # self.answer.setAlignment(QtCore.Qt.AlignLeading | QtCore.Qt.AlignLeft | QtCore.Qt.AlignTop)
        self.answer.setStyleSheet('''
                                    QScrollArea {
                                            background-color: #FFFFFF;
                                            margin-right: 0px;
                                            width: 500%;
                                    }
                                ''')

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
                                    margin: 20px 40px;
                                    width: 500%;
                            }
                        ''')

        self.downloadFile = QPushButton("Скачать проверенный\nфайл", self)
        self.downloadFile.setMinimumSize(350, 80)
        self.downloadFile.setStyleSheet('''
                                    QPushButton {
                                            font-weight: 700;
                                            background-color: #E9E9E9;
                                            font-family: 'Aleo';
                                            font-size: 20px;
                                            color: #000000;
                                            border: 3px solid #FFFFFF;
                                            border-radius: 36px;
                                            margin: 30px 50px 0px 50px;
                                            padding: 10px 0px;
                                            width: 70%; height: 50%;
                                        }
                                    ''')

        # ДОБАВЛЕНИЕ ЭЛЕМЕНТОВ В ГРИД #
        grid = QGridLayout()
        grid.setSpacing(0)
        self.setLayout(grid)
        # grid.setSpacing(10)
        grid.setContentsMargins(62, 0, 0, 0)
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
        grid.addWidget(self.confirm_button, 8, 0, 1, 2)  # QPushButton('Подтвердить настройки')
        # grid.addWidget(self.pickFile, 9, 0)  # QLabel('Выберите файл (docx):')
        grid.addWidget(self.pickFileButton, 9, 0, 1, 2)  # QPushButton("Выбрать файл")
        grid.addWidget(self.filePicked, 10, 0, 1, 2)  # QLabel("", self) выбранный файл
        grid.addWidget(self.checkFile, 11, 0, 1, 2)  # QPushButton('Проверить файл')
        grid.addWidget(self.answer, 1, 2, 11, 1)  # QScrollArea(self)
        grid.addWidget(self.plain_text, 1, 2, 9, 1)  # QPlainTextEdit() поле с ответом
        grid.addWidget(self.downloadFile, 10, 2, 2, 1)  # QPushButton("Скачать проверенный файл")
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


    def open_second_window(self):
        from secondWindow import SecondWindow
        if not self.second_window:
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
        with open('./files/gost/' + self.filename_settings + '.json', "w", encoding='utf-8') as json_file:
            json.dump(data, json_file, indent=4, ensure_ascii=False)

    def changeIndentLabel(self, text):
        self.enterIndentLabel.setText(text + ' см')

    def changeLineSpaceLabel(self, text):
        self.enterLineSpaceLabel.setText(text + ' см')

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


if __name__ == '__main__':
    app = QApplication(sys.argv)
    # app.setStyleSheet(stream.readAll())
    main_window = MainWindow()
    main_window.show()
    sys.exit(app.exec_())