import json
import openpyxl
import os.path
from datetime import date
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QMessageBox, QMainWindow
from openpyxl.styles import Alignment, Border, Side
from PyQt5 import QtWidgets
from selenium import webdriver
from bs4 import BeautifulSoup
import time
from selenium.webdriver.firefox.options import Options as FirefoxOptions
from selenium.webdriver.firefox.options import Options

class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(323, 170)
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap("232968.ico"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        MainWindow.setWindowIcon(icon)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.gridLayout = QtWidgets.QGridLayout(self.centralwidget)
        self.gridLayout.setObjectName("gridLayout")
        self.checkBox = QtWidgets.QCheckBox(self.centralwidget)
        self.checkBox.setAutoFillBackground(False)
        self.checkBox.setChecked(False)
        self.checkBox.setObjectName("checkBox")
        self.gridLayout.addWidget(self.checkBox, 1, 0, 1, 1)
        self.label = QtWidgets.QLabel(self.centralwidget)
        self.label.setObjectName("label")
        self.gridLayout.addWidget(self.label, 3, 0, 1, 1)
        self.label_2 = QtWidgets.QLabel(self.centralwidget)
        font = QtGui.QFont()
        font.setStyleStrategy(QtGui.QFont.PreferDefault)
        self.label_2.setFont(font)
        self.label_2.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.label_2.setAutoFillBackground(False)
        self.label_2.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_2.setObjectName("label_2")
        self.gridLayout.addWidget(self.label_2, 3, 1, 1, 1)
        self.pushButton = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton.setObjectName("pushButton")
        self.gridLayout.addWidget(self.pushButton, 2, 0, 1, 2)
        self.plainTextEdit = QtWidgets.QPlainTextEdit(self.centralwidget)
        self.plainTextEdit.setAutoFillBackground(False)
        self.plainTextEdit.setObjectName("plainTextEdit")
        self.gridLayout.addWidget(self.plainTextEdit, 0, 0, 1, 2)
        self.label_3 = QtWidgets.QLabel(self.centralwidget)
        self.label_3.setAutoFillBackground(False)
        self.label_3.setTextFormat(QtCore.Qt.AutoText)
        self.label_3.setScaledContents(False)
        self.label_3.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_3.setWordWrap(False)
        self.label_3.setIndent(-1)
        self.label_3.setOpenExternalLinks(True)
        self.label_3.setTextInteractionFlags(QtCore.Qt.LinksAccessibleByMouse|QtCore.Qt.TextSelectableByMouse)
        self.label_3.setObjectName("label_3")

        self.gridLayout.addWidget(self.label_3, 1, 1, 1, 1)
        MainWindow.setCentralWidget(self.centralwidget)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)
        self.pushButton.clicked.connect(self.applicat)
    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "Парсер заявок"))
        self.checkBox.setText(_translate("MainWindow", "Сохранить копию"))
        self.label.setText(_translate("MainWindow", "Rols парсер заявок"))
        self.label_2.setText(_translate("MainWindow", "Версия 1.2"))
        self.pushButton.setText(_translate("MainWindow", "Распарсить"))
        self.label_3.setText('<a style="color:red;" href="https://hd.crystals.ru/api/Claim/GetUserActionHistory"> Ссылка на историю </a>')
    
    def applicat(self):
        _translate = QtCore.QCoreApplication.translate
        if self.plainTextEdit.toPlainText() == '':
            self.Text_messagebox()
        else:
            try:
                #
                    #---------------------------------------------------
                options = Options()
                #options.add_argument("--headless") # Настройка отвечающая за фоновую активность
                options.binary_location = r'C:\Program Files\Mozilla Firefox\firefox.exe'
                browser = webdriver.Firefox(executable_path=r'D:\Proect_python\geckodriver.exe', options=options)
                #---------------------------------------------------
                home_page = "https://hd.crystals.ru/"
                schedule_page = "https://hd.crystals.ru/api/Claim/GetUserActionHistory"
                browser.get(home_page)
                time.sleep(0.5)
                login_input = browser.find_element_by_id('claim_login')
                password_input = browser.find_element_by_id('claim_pass')
                login_input.send_keys('Мешков Анатолий')
                password_input.send_keys('T+Ye&Z4(')
                submit_button = browser.find_element_by_xpath("//input[@type='submit']")
                submit_button.click()
                time.sleep(0.5)
                browser.get(schedule_page)
                html = browser.page_source
                soup = BeautifulSoup(html)
                dannie = soup.text
                #
                self.plainTextEdit.setPlainText(_translate('MainWindow', dannie))
                print(self.plainTextEdit.toPlainText())
                #
                data = json.loads(self.plainTextEdit.toPlainText()[self.plainTextEdit.toPlainText().find('{')-1:self.plainTextEdit.toPlainText().rfind('}') + 2])
                #___________________________________________
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
                        if "Fix Price" in name[2]:
                            print(name)
                        if name[2][:6] == 'Верный':
                            name[2] = 'Верный'
                        if name[2] == "О'КЕЙ":
                            name[2] = 'Окей'
                        if name[1] == 'Матюшевский Александр':
                            name[1] = 'Матюшевский'
                        elif name[1] == 'Шешуков Денис':
                            name[1] = 'Шешуков'
                        elif name[1] == 'Грошев Владислав':
                            name[1] = 'Грошев'
                        elif name[1] == 'Каверин Евгений':
                            name[1] = 'Каверин'
                        elif name[1] == 'Мешков Анатолий':
                            name[1] = 'Мешков'
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

    def Error_count_messagebox(self):
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Information)
        msg.setText(f'Завиксировано одновременное принятия нескольких заявок пример отдной из них:{self.nan}')
        msg.setWindowTitle("Справка")
        msg.setStandardButtons(QMessageBox.Ok)
        retval = msg.exec_()

    def show_info_messagebox(self):
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Information)
        # setting message for Message Box
        msg.setText("Закройте файл xlsx с заявками или удалите!")
        # setting Message box window title
        msg.setWindowTitle("Справка")
        # declaring buttons on Message Box
        msg.setStandardButtons(QMessageBox.Ok | QMessageBox.Cancel)
        # start the app
        retval = msg.exec_()

    def Text_messagebox(self):
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Warning)
        msg.setText("Полле ввода пусто!")
        msg.setWindowTitle("Справка")
        msg.setStandardButtons(QMessageBox.Ok)
        retval = msg.exec_()

    def Text_govna_messagebox(self):
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Information)
        msg.setText('''Введена какая-то херь!
Попробуй ещё раз)''')
        msg.setWindowTitle("Справка")
        msg.setStandardButtons(QMessageBox.Ok)
        retval = msg.exec_()

    def App_good_messagebox(self):
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Information)
        msg.setText("Файл example.xlsx успешно сформировался")
        msg.setWindowTitle("Справка")
        msg.setStandardButtons(QMessageBox.Ok)
        retval = msg.exec_()


if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())
