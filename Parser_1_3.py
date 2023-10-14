import json
import openpyxl
import os.path
import telebot
from datetime import date
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QMessageBox, QMainWindow
from openpyxl.styles import Alignment, Border, Side
import time


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(406, 176)
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap("232968.ico"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        MainWindow.setWindowIcon(icon)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.gridLayout = QtWidgets.QGridLayout(self.centralwidget)
        self.gridLayout.setObjectName("gridLayout")
        self.verticalLayout = QtWidgets.QVBoxLayout()
        self.verticalLayout.setObjectName("verticalLayout")
        self.plainTextEdit = QtWidgets.QPlainTextEdit(self.centralwidget)
        self.plainTextEdit.setAutoFillBackground(False)
        self.plainTextEdit.setObjectName("plainTextEdit")
        self.verticalLayout.addWidget(self.plainTextEdit)
        self.horizontalLayout_2 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_2.setObjectName("horizontalLayout_2")
        self.checkBox = QtWidgets.QCheckBox(self.centralwidget)
        self.checkBox.setAutoFillBackground(False)
        self.checkBox.setChecked(False)
        self.checkBox.setObjectName("checkBox")
        self.horizontalLayout_2.addWidget(self.checkBox)
        self.label_3 = QtWidgets.QLabel(self.centralwidget)
        self.label_3.setEnabled(True)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.label_3.sizePolicy().hasHeightForWidth())
        self.label_3.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setFamily("MS Shell Dlg 2")
        font.setPointSize(8)
        font.setBold(False)
        font.setItalic(False)
        font.setWeight(50)
        font.setKerning(False)
        self.label_3.setFont(font)
        self.label_3.setMouseTracking(False)
        self.label_3.setTabletTracking(False)
        self.label_3.setFocusPolicy(QtCore.Qt.WheelFocus)
        self.label_3.setContextMenuPolicy(QtCore.Qt.DefaultContextMenu)
        self.label_3.setAcceptDrops(False)
        self.label_3.setAutoFillBackground(True)
        self.label_3.setStyleSheet("font: 8pt \"Times New Roman\";\n"
"font: 8pt \"MS Shell Dlg 2\";\n"
"")
        self.label_3.setInputMethodHints(QtCore.Qt.ImhExclusiveInputMask)
        self.label_3.setFrameShape(QtWidgets.QFrame.NoFrame)
        self.label_3.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.label_3.setTextFormat(QtCore.Qt.AutoText)
        self.label_3.setScaledContents(False)
        self.label_3.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_3.setWordWrap(False)
        self.label_3.setIndent(-1)
        self.label_3.setOpenExternalLinks(True)
        self.label_3.setTextInteractionFlags(QtCore.Qt.LinksAccessibleByKeyboard|QtCore.Qt.LinksAccessibleByMouse|QtCore.Qt.TextBrowserInteraction|QtCore.Qt.TextSelectableByKeyboard|QtCore.Qt.TextSelectableByMouse)
        self.label_3.setObjectName("label_3")
        self.horizontalLayout_2.addWidget(self.label_3)
        self.verticalLayout.addLayout(self.horizontalLayout_2)
        self.pushButton = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton.setObjectName("pushButton")
        self.verticalLayout.addWidget(self.pushButton)
        self.horizontalLayout = QtWidgets.QHBoxLayout()
        self.horizontalLayout.setObjectName("horizontalLayout")
        self.label = QtWidgets.QLabel(self.centralwidget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.label.sizePolicy().hasHeightForWidth())
        self.label.setSizePolicy(sizePolicy)
        self.label.setObjectName("label")
        self.horizontalLayout.addWidget(self.label)
        self.pushButton_2 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_2.setEnabled(True)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.pushButton_2.sizePolicy().hasHeightForWidth())
        self.pushButton_2.setSizePolicy(sizePolicy)
        self.pushButton_2.setMaximumSize(QtCore.QSize(16700000, 16700000))
        self.pushButton_2.setAutoDefault(False)
        self.pushButton_2.setObjectName("pushButton_2")
        self.horizontalLayout.addWidget(self.pushButton_2)
        self.label_2 = QtWidgets.QLabel(self.centralwidget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.label_2.sizePolicy().hasHeightForWidth())
        self.label_2.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setStyleStrategy(QtGui.QFont.PreferDefault)
        self.label_2.setFont(font)
        self.label_2.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.label_2.setAutoFillBackground(False)
        self.label_2.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_2.setObjectName("label_2")
        self.horizontalLayout.addWidget(self.label_2)
        self.verticalLayout.addLayout(self.horizontalLayout)
        self.gridLayout.addLayout(self.verticalLayout, 0, 0, 1, 1)
        MainWindow.setCentralWidget(self.centralwidget)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)
        self.pushButton.clicked.connect(self.applicat)
        self.pushButton_2.clicked.connect(self.sendung)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "Парсер заявок"))
        self.checkBox.setText(_translate("MainWindow", "Сохранить копию"))
        self.label_3.setText('<a style="color:red;" href="https://hd.crystals.ru/api/Claim/GetUserActionHistory"> Ссылка на историю </a>')
        self.pushButton.setText(_translate("MainWindow", "Распарсить"))
        self.label.setText(_translate("MainWindow", "Rols парсер заявок"))
        self.label_2.setText(_translate("MainWindow", "Версия 1.3"))
        self.pushButton_2.setText(_translate("MainWindow", "SOS"))

    def bearbeiter_text(self):
        data = json.loads(self.plainTextEdit.toPlainText()[self.plainTextEdit.toPlainText().find('{')-1:self.plainTextEdit.toPlainText().rfind('}') + 2])
        return data

    def applicat(self):
        _translate = QtCore.QCoreApplication.translate
        if self.plainTextEdit.toPlainText() == '':
            self.Text_messagebox()
        else:
            try:
                data = self.bearbeiter_text()
                wb = openpyxl.Workbook()
                wb.create_sheet(title = 'Первый лист', index = 0)
                sheet = wb['Первый лист']
                Completed_applications, applications, telo, applications_telo = [], [], [] ,[]
                center_align = Alignment(horizontal='right', vertical='center')
                Border_siz = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
                #Список заявок
                for good in data:
                    if 'Выполнить' in good['State'] and 'Выполнена' in good['ClaimState'] or '3 линия техподдержки' in good['Depart']:
                        if good['ClaimID'] in Completed_applications:
                            continue
                        else:
                            Completed_applications.append(good['ClaimID'])
                            # Поиск данных выполненой заявки
                            telo.append([good['Dispatcher'], good['ClientCompName'], good['ClaimID'], ' ', good['StartDate'], good['Depart']])
                # Поиск взятия выполненой заявки
                for good_pri in data:
                    if good_pri['ClaimState'] in 'Принята к исполнению' and good_pri['ClaimID'] in Completed_applications:
                        if good_pri['StartDate'] in applications:
                            continue
                        else:
                            applications.append(good_pri['StartDate'])
                            Completed_applications.remove(good_pri['ClaimID'])
                #------Обходное решение
                if len(applications) < len(telo):
                    self.nan = Completed_applications
                    self.Error_count_messagebox()
                    for i in range(len(telo)-len(applications)):
                        applications.append(applications[-1])
                #Обънединение списков а так же привидение в необходимый вид и запись в excel      
                for i in range(len(telo)):
                    applications_telo += [[applications[::-1][i]] + telo[::-1][i]]
                    for name in applications_telo:
                        if "Прайс" in name[2]: name[2] = 'Fix Price'
                        if "Верный" in name[2]:  name[2] = 'Верный'
                        if "О'КЕЙ" in name[2]: name[2] = 'Окей'
                        if "ВИНЛАБ" in name[2]: name[2] = 'Винлаб'
                        if "ВИКТОРИЯ" in name[2]: name[2] = 'Виктория'
                        if "Азбука Вкуса" in name[2]: name[2] = 'Азбука Вкуса'
                        if name[1] == 'Матюшевский Александр': name[1] = 'Матюшевский'
                        elif name[1] == 'Шешуков Денис': name[1] = 'Шешуков'
                        elif name[1] == 'Грошев Владислав': name[1] = 'Грошев'
                        elif name[1] == 'Каверин Евгений': name[1] = 'Каверин'
                        elif name[1] == 'Мешков Анатолий': name[1] = 'Мешков'
                        elif name[1] == 'Захарихин Евгений': name[1] = 'Захарихин'
                        elif name[1] == 'Иванов Никита': name[1] = 'Иванов'
                        elif name[1] == 'Мурзин Анатолий': name[1] = 'Мурзин'
                        elif name[1] == 'Невский Никита': name[1] = 'Невский'
                        elif name[1] == 'Степанов Илья': name[1] = 'Степанов'
                        elif name[1] == 'Черничкин Павел': name[1] = 'Черничкин'
                        if name[6] != '3 линия техподдержки':
                            name[6] = ' '
                    sheet.append(applications_telo[i])
                #Сохранение в необходимомм формате
                try:
                    for i in 'ABCDEF':
                        for ji in sheet[i+'1:'+ i + str(len(applications_telo))]:
                            ji[0].alignment = center_align
                            ji[0].border = Border_siz
                    sheet.column_dimensions['A'].width = 20
                    sheet.column_dimensions['B'].width = 15
                    sheet.column_dimensions['C'].width = 15
                    sheet.column_dimensions['E'].width = 20
                    sheet.column_dimensions['F'].width = 20
                    name_xlsx = "example_" + str(date.today()) + ".xlsx"
                    if True == self.checkBox.isChecked():
                        for i in range(10**2):
                            if True != os.path.isfile(name_xlsx):
                                wb.save(name_xlsx)
                                self.App_good_messagebox()
                                break
                            else:
                                name_xlsx = name_xlsx[:8] + str(date.today()) + '_' + str(i) + name_xlsx[-5:]                  
                    else:
                        wb.save(name_xlsx)
                        self.App_good_messagebox()
                except:
                    self.show_info_messagebox()
                #___________________________________________
            except:
                self.Text_govna_messagebox()
    
    def sendung(self):
        token = "5140763739:AAFksKRlBdYAhNm7TSOdkZBbdPaT3HEz7oo"
        bot = telebot.TeleBot(token)
        if self.plainTextEdit.toPlainText() == '':
            self.Text_messagebox()
        else:
            try:
                with open("Bericht.txt", "w") as file:
                    file.write(self.plainTextEdit.toPlainText())
                file.close()
                file = open("Bericht.txt","rb")
                bot.send_document('1007004173', document=file, caption='Проблемный файл из парсера c ошибкой!')
                file.close()
                os.remove("Bericht.txt")
                self.Good_sending()
            except :
                self.Error_telebot_messagebox()

    def Text_messagebox(self):
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Warning)
        msg.setText("Полле ввода пусто!")
        msg.setWindowTitle("Справка")
        msg.setStandardButtons(QMessageBox.Ok)
        msg.exec_()

    def Error_telebot_messagebox(self):
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Warning)
        msg.setText("Данне не отправленны создателю!")
        msg.setWindowTitle("ERROR")
        msg.setStandardButtons(QMessageBox.Ok)
        msg.exec_()

    def Error_count_messagebox(self):
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Information)
        msg.setText(f'Завиксировано одновременное принятия нескольких заявок пример отдной из них:{self.nan}')
        msg.setWindowTitle("Справка")
        msg.setStandardButtons(QMessageBox.Ok)
        msg.exec_()

    def show_info_messagebox(self):
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Information)
        msg.setText("Закройте файл xlsx с заявками или удалите!")
        msg.setWindowTitle("Справка")
        msg.setStandardButtons(QMessageBox.Ok)
        msg.exec_()

    def Text_govna_messagebox(self):
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Information)
        msg.setText('''Введена какая-то херь!
Попробуй ещё раз)''')
        msg.setWindowTitle("Справка")
        msg.setStandardButtons(QMessageBox.Ok)
        msg.exec_()

    def App_good_messagebox(self):
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Information)
        msg.setText("Файл example.xlsx успешно сформировался")
        msg.setWindowTitle("Справка")
        msg.setStandardButtons(QMessageBox.Ok)
        msg.exec_()

    def Good_sending(self):
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Information)
        msg.setText("Отчёт сформирован и отправлен создателю")
        msg.setWindowTitle("Справка")
        msg.setStandardButtons(QMessageBox.Ok)
        msg.exec_()

if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())
