import sys
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5 import QtCore
import lib.Create_PLC_OMX as Create_PLC_OMX
import lib.Create_Type_OMX as Create_Type_OMX
import lib.Create_AttrMap_OMX as Create_AttrMap_OMX
import time
import datetime
import subprocess


def click1(file_name, sheet_name, edit):
    edit.append(f'{datetime.datetime.now()}')
    edit.append("Start generation PLC")
    edit.append(f'File: {file_name}, sheet: {sheet_name}')
    try:
        Create_PLC_OMX.main_create(file_name, sheet_name)
        edit.append("Finish generation PLC\n")
    except Exception as e:
        edit.append(f"Error generation: {e}\n")
    edit.append("*******************")
    cursor = edit.textCursor()
    cursor.movePosition(QTextCursor.End) 
    edit.setTextCursor(cursor)

def click2(file_name, sheet_name, name_type, namespace, edit):
    edit.append(f'{datetime.datetime.now()}')
    edit.append("Start generation Type")
    edit.append(f'File: {file_name}, sheet: {sheet_name}')
    try:
        Create_Type_OMX.main_create(file_name, sheet_name, name_type, namespace)
        edit.append("Finish generation Type\n")
    except Exception as e:
        edit.append(f"Error generation: {e}\n")
    edit.append("*******************")
    cursor = edit.textCursor()
    cursor.movePosition(QTextCursor.End) 
    edit.setTextCursor(cursor)

def click3(file_name, sheet_name, name_element, parametr, edit):
    edit.append(f'{datetime.datetime.now()}')
    edit.append("Start generation Attribute")
    edit.append(f'File: {file_name}, sheet: {sheet_name}')
    try:
        Create_AttrMap_OMX.main_create(file_name, sheet_name, name_element, parametr)
        edit.append("Finish generation Attribute\n")
    except Exception as e:
        edit.append(f"Error generation: {e}\n")
    edit.append("*******************")
    cursor = edit.textCursor()
    cursor.movePosition(QTextCursor.End) 
    edit.setTextCursor(cursor)

def click11(edit):

    edit.append("Создание omx-файла для импорта объектов в ПЛК:\n - Имя импортируемого файла \n - Имя импортируемого листа")
    edit.append("\nЛист должен содержать следующие столбцы:\n - CodePLC - наименование элемента в проекте АСУТП\n"+
                " - Description - описание элемента\n - BaseType - тип элемента (требуется полный путь от папки Types)\n")
    edit.append("*******************")
    cursor = edit.textCursor()
    cursor.movePosition(QTextCursor.End) 
    edit.setTextCursor(cursor)

def click21(edit):

    
    edit.append("Создание omx-файла для импорта Типов:\n - Имя импортируемого файла \n - Имя импортируемого листа\n - Имя типа"+
                "\n - Имя пространства (папка для хранения типа в Types)")
    edit.append("\nЛист должен содержать следующие столбцы:\n - Name - наименование элемента в типе\n - Description - описание элемента"+
                "\n - Unit - ед. измерения элемента\n - Direction - направление элемента (in, out, in-out)\n - Type - тип элемента\n")
    
    edit.append("*******************")
    cursor = edit.textCursor()
    cursor.movePosition(QTextCursor.End) 
    edit.setTextCursor(cursor)

def click31(edit):
    
    edit.append("Создание omx-файла для импорта списка Атрибутов:\n - Имя импортируемого файла \n - Имя импортируемого листа\n - Имя итогового файла xml со значениями атрибутов"+
                "\n - Тип Атрибута (например, Unit - ед. измерения, Description - описание)")
    edit.append("\nЛист должен содержать следующие столбцы:\n - FullName - полное наименование элемента\n - Description/Unit/др.параметр - тип атрибута\n")
    
    edit.append("*******************")
    cursor = edit.textCursor()
    cursor.movePosition(QTextCursor.End) 
    edit.setTextCursor(cursor)

def main():
    
    app = QApplication(sys.argv)
    win = QMainWindow()
    win.setFixedSize(600,700)
    win.setWindowTitle("Создание файла импорта")
    win.setWindowIcon(QIcon('file/icon.png'))

    title = QLabel("", win)
    title.setFont(QFont('Arial', 16))
    title.setAlignment(QtCore.Qt.AlignCenter)
    title.resize(560, 30)
    title.move(20,50)
    
    frame1 = QFrame(win)
    frame1.setFrameShape(QFrame.StyledPanel)
    frame1.setGeometry(10, 100, 580, 280)
    frame1.setStyleSheet(f"QFrame{{border: 0px;}}")

    frame2 = QFrame(win)
    frame2.setFrameShape(QFrame.StyledPanel)
    frame2.setGeometry(10, 100, 580, 280)
    frame2.setStyleSheet(f"QFrame{{border: 0px;}}")

    frame3 = QFrame(win)
    frame3.setFrameShape(QFrame.StyledPanel)
    frame3.setGeometry(10, 100, 580, 280)
    frame3.setStyleSheet(f"QFrame{{border: 0px;}}")

    frame1.hide()
    frame2.hide()
    frame3.hide()

    button_PLC = QAction("Генерация объектов ПЛК", win)
    button_PLC.triggered.connect(lambda: frame1.show())
    button_PLC.triggered.connect(lambda: frame2.hide())
    button_PLC.triggered.connect(lambda: frame3.hide())
    button_PLC.triggered.connect(lambda: title.setText("Генерация объектов ПЛК"))

    button_type = QAction("Генерация типов", win)
    button_type.triggered.connect(lambda: frame2.show())
    button_type.triggered.connect(lambda: frame1.hide())
    button_type.triggered.connect(lambda: frame3.hide())
    button_type.triggered.connect(lambda: title.setText("Генерация типов"))

    button_atr = QAction("Генерация списка атрибутов", win)
    button_atr.triggered.connect(lambda: frame3.show())
    button_atr.triggered.connect(lambda: frame1.hide())
    button_atr.triggered.connect(lambda: frame2.hide())
    button_atr.triggered.connect(lambda: title.setText("Генерация списка атрибутов"))

    button_close = QAction("Закрыть программу", win)
    button_close.triggered.connect(lambda: win.close())

    button_clear = QAction("Очистка логов", win)
    button_clear.triggered.connect(lambda: logger.clear())

    button_faq = QAction("FAQ", win)
    button_faq.triggered.connect(lambda: subprocess.run(["notepad","ReadMe.txt"]))

    menu = win.menuBar()
    file_menu = menu.addMenu("Меню")
    file_menu.addAction(button_clear)
    file_menu.addAction(button_faq)
    file_menu.addAction(button_close)
    
    generartion_menu = menu.addMenu("Выбор объектов генерации")
    generartion_menu.addAction(button_PLC)
    generartion_menu.addAction(button_type)
    generartion_menu.addAction(button_atr)
    
    logger = QTextEdit("", win)
    logger.setAlignment(QtCore.Qt.AlignLeft)
    logger.setFont(QFont('Arial', 10))
    logger.move(20, 400)
    logger.resize(560,280)
    logger.setReadOnly(True)
        
    PLC_label1= QLabel("Имя импорт. файла", frame1)
    PLC_label1.setFont(QFont('Arial', 12))
    PLC_label1.resize(260, 30)
    PLC_label1.move(20,20)

    PLC_label2= QLabel("Имя импорт. листа", frame1)
    PLC_label2.setFont(QFont('Arial', 12))
    PLC_label2.resize(260, 30)
    PLC_label2.move(20,60)

    PLC_textbox1 = QLineEdit("Input.xlsx", frame1)
    PLC_textbox1.setAlignment(QtCore.Qt.AlignLeft)
    PLC_textbox1.setFont(QFont('Arial', 12))
    PLC_textbox1.move(300, 20)
    PLC_textbox1.resize(260,30)

    PLC_textbox2 = QLineEdit("CreateOMX", frame1)
    PLC_textbox2.setAlignment(QtCore.Qt.AlignLeft)
    PLC_textbox2.setFont(QFont('Arial', 12))
    PLC_textbox2.move(300, 60)
    PLC_textbox2.resize(260,30)

    PLC_button1 = QPushButton("Генерация", frame1)
    PLC_button1.move(20,220)
    PLC_button1.setFont(QFont('Arial', 12))
    PLC_button1.resize(220, 40)
    PLC_button1.clicked.connect(lambda: click1(PLC_textbox1.text(), PLC_textbox2.text(), logger))

    PLC_button2 = QPushButton("Очистка ячеек", frame1)
    PLC_button2.move(270,220)
    PLC_button2.setFont(QFont('Arial', 12))
    PLC_button2.resize(220, 40)
    PLC_button2.clicked.connect(PLC_textbox1.clear)
    PLC_button2.clicked.connect(PLC_textbox2.clear)

    PLC_button3 = QPushButton("?", frame1)
    PLC_button3.move(520,220)
    PLC_button3.setFont(QFont('Arial', 12))
    PLC_button3.resize(40, 40)
    PLC_button3.clicked.connect(lambda: click11(logger))

    type_label1= QLabel("Имя импорт. файла", frame2)
    type_label1.setFont(QFont('Arial', 12))
    type_label1.resize(260, 30)
    type_label1.move(20,20)

    type_label2= QLabel("Имя импорт. листа", frame2)
    type_label2.setFont(QFont('Arial', 12))
    type_label2.resize(260, 30)
    type_label2.move(20,60)

    type_label3= QLabel("Имя типа", frame2)
    type_label3.setFont(QFont('Arial', 12))
    type_label3.resize(260, 30)
    type_label3.move(20,100)

    type_label4= QLabel("Имя пространства", frame2)
    type_label4.setFont(QFont('Arial', 12))
    type_label4.resize(260, 30)
    type_label4.move(20,140)

    type_textbox1 = QLineEdit("Input.xlsx", frame2)
    type_textbox1.setAlignment(QtCore.Qt.AlignLeft)
    type_textbox1.setFont(QFont('Arial', 12))
    type_textbox1.move(300, 20)
    type_textbox1.resize(260,30)

    type_textbox2 = QLineEdit("CreateType", frame2)
    type_textbox2.setAlignment(QtCore.Qt.AlignLeft)
    type_textbox2.setFont(QFont('Arial', 12))
    type_textbox2.move(300, 60)
    type_textbox2.resize(260,30)

    type_textbox3 = QLineEdit("TypeName", frame2)
    type_textbox3.setAlignment(QtCore.Qt.AlignLeft)
    type_textbox3.setFont(QFont('Arial', 12))
    type_textbox3.move(300, 100)
    type_textbox3.resize(260,30)

    type_textbox4 = QLineEdit("TypeNamespace", frame2)
    type_textbox4.setAlignment(QtCore.Qt.AlignLeft)
    type_textbox4.setFont(QFont('Arial', 12))
    type_textbox4.move(300, 140)
    type_textbox4.resize(260,30)

    type_button1 = QPushButton("Генерация", frame2)
    type_button1.move(20,220)
    type_button1.setFont(QFont('Arial', 12))
    type_button1.resize(220, 40)
    type_button1.clicked.connect(lambda: click2(type_textbox1.text(), type_textbox2.text(), type_textbox3.text(), type_textbox4.text(),logger))

    type_button2 = QPushButton("Очистка ячеек", frame2)
    type_button2.move(270,220)
    type_button2.setFont(QFont('Arial', 12))
    type_button2.resize(220, 40)
    type_button2.clicked.connect(type_textbox1.clear)
    type_button2.clicked.connect(type_textbox2.clear)
    type_button2.clicked.connect(type_textbox3.clear)
    type_button2.clicked.connect(type_textbox4.clear)

    type_button3 = QPushButton("?", frame2)
    type_button3.move(520,220)
    type_button3.setFont(QFont('Arial', 12))
    type_button3.resize(40, 40)
    type_button3.clicked.connect(lambda: click21(logger))

    atr_label1= QLabel("Имя импорт. файла", frame3)
    atr_label1.setFont(QFont('Arial', 12))
    atr_label1.resize(260, 30)
    atr_label1.move(20,20)

    atr_label2= QLabel("Имя импорт. листа", frame3)
    atr_label2.setFont(QFont('Arial', 12))
    atr_label2.resize(260, 30)
    atr_label2.move(20,60)

    atr_label3= QLabel("Имя итогового файла", frame3)
    atr_label3.setFont(QFont('Arial', 12))
    atr_label3.resize(260, 30)
    atr_label3.move(20,100)

    atr_label4= QLabel("Тип атрибутов", frame3)
    atr_label4.setFont(QFont('Arial', 12))
    atr_label4.resize(260, 30)
    atr_label4.move(20,140)

    atr_textbox1 = QLineEdit("Input.xlsx", frame3)
    atr_textbox1.setAlignment(QtCore.Qt.AlignLeft)
    atr_textbox1.setFont(QFont('Arial', 12))
    atr_textbox1.move(300, 20)
    atr_textbox1.resize(260,30)

    atr_textbox2 = QLineEdit("CreateType", frame3)
    atr_textbox2.setAlignment(QtCore.Qt.AlignLeft)
    atr_textbox2.setFont(QFont('Arial', 12))
    atr_textbox2.move(300, 60)
    atr_textbox2.resize(260,30)

    atr_textbox3 = QLineEdit("Attrib", frame3)
    atr_textbox3.setAlignment(QtCore.Qt.AlignLeft)
    atr_textbox3.setFont(QFont('Arial', 12))
    atr_textbox3.move(300, 100)
    atr_textbox3.resize(260,30)

    atr_textbox4 = QComboBox(frame3)
    paramlist = ["Unit", "Description"]
    atr_textbox4.addItems(paramlist)
    atr_textbox4.setFont(QFont('Arial', 12))
    atr_textbox4.move(300, 140)
    atr_textbox4.resize(260,30)
    atr_textbox4.setEditable(False)
    atr_textbox4.setInsertPolicy(QComboBox.NoInsert)
    
    atr_button1 = QPushButton("Генерация", frame3)
    atr_button1.move(20,220)
    atr_button1.setFont(QFont('Arial', 12))
    atr_button1.resize(220, 40)
    atr_button1.clicked.connect(lambda: click3(atr_textbox1.text(), atr_textbox2.text(), atr_textbox3.text(), atr_textbox4.currentText(),logger))

    atr_button2 = QPushButton("Очистка ячеек", frame3)
    atr_button2.move(270,220)
    atr_button2.setFont(QFont('Arial', 12))
    atr_button2.resize(220, 40)
    atr_button2.clicked.connect(atr_textbox1.clear)
    atr_button2.clicked.connect(atr_textbox2.clear)
    atr_button2.clicked.connect(atr_textbox3.clear)

    atr_button3 = QPushButton("?", frame3)
    atr_button3.move(520,220)
    atr_button3.setFont(QFont('Arial', 12))
    atr_button3.resize(40, 40)
    atr_button3.clicked.connect(lambda: click31(logger))
    
    win.show()
    sys.exit(app.exec_())

    
main()
