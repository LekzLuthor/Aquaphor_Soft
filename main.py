import json
import sys
import datetime
import os
import openpyxl
import pprint
import os

from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QApplication, QMainWindow, QTextEdit, QMessageBox, QTimeEdit
from PyQt5 import uic
import smtplib
from email.message import EmailMessage


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        uic.loadUi('main_window.ui', self)

        # блок данных программы
        self.EMAIL_SENDER = "koshanskiy00@mail.ru"
        self.EMAIL_SENDER_PASS = ""
        self.settings = {}
        self.excel_files_names = []  # список названий файлов
        self.database = {}  # база данных сформированная из файлов
        self.equipment_report = {}  # оборудование с просроченной датой калибровки
        self.emails = []

        # привязка кнопок + редактура полей
        self.generateReportButton.clicked.connect(self.send_report)
        self.saveChangesButton.clicked.connect(self.save_changes)
        self.emailSetForm.setPlaceholderText("example@aquaphor.com")
        self.pathwaySetForm.setPlaceholderText("C:/path/path with your Excel Files")

        # кнопка для тестов
        self.test.clicked.connect(self.create_report)

        # блок логики программы при запуске
        with open("sours/settings.json", "r") as file:  # достаёт настройки из json файла
            self.settings = json.load(file)
            for email in self.settings["emails"]:
                self.emails.append(email)

        if self.check_pathway():
            self.load_database()
            print("done")
        else:
            pass

    def save_changes(self):
        email = self.emailSetForm.toPlainText()
        if email != "":
            if self.check_email(email):
                if email not in self.settings["emails"]:
                    self.settings["emails"].append(email)
                    self.emailStatusBar.setText('mail added to list')
                else:
                    self.emailStatusBar.setText("you haven't entered a new email")
            else:
                pass
        else:
            self.emailStatusBar.setText("you haven't entered a new email")

        if self.reportTime.time().toPyTime() != "00:00:00":
            self.settings["report_time"] = str(self.reportTime.time().toPyTime())
            self.timeStatusBar.setText('report time saved')
        elif self.reportTime.time().toPyTime() == self.settings['report_time']:
            self.timeStatusBar.setText("you haven't entered a new time")
        else:
            self.timeStatusBar.setText("you haven't entered a new time")

        # Получение файлового пути и кэтч ошибки с неправильным (\) символом
        try:
            pathway = self.pathwaySetForm.toPlainText()
            try:
                with open(f'{pathway}/Your Excel Files Will Be here', 'w') as f:
                    f.writelines('Checking the correctness of path way')
                if os.path.isfile(f'{pathway}/Your Excel Files Will Be here'):
                    os.remove(f'{pathway}/Your Excel Files Will Be here')

            except OSError:
                msg = QMessageBox()
                msg.setIcon(QMessageBox.Warning)
                msg.setText("Error: WRONG PATH WAY")
                msg.setWindowTitle("Warning")
                retval = msg.exec_()
                self.pathwayStatusBar.setText('wrong path way')

            if pathway != '':
                self.settings["pathway"] = pathway
                self.pathwayStatusBar.setText('pathway saved')

        except SyntaxError:
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Warning)
            msg.setText("Error: WRONG PATH WAY")
            msg.setWindowTitle("Warning")
            retval = msg.exec_()
            self.pathwayStatusBar.setText('wrong path way')

        with open("sours/settings.json", "w") as file:  # сохраняет настройки в json файл
            json.dump(self.settings, file)

    @staticmethod
    def check_email(email):  # проверка правильности ввода почты
        if email and (
                "@aquaphor.com" in email or "@mail.ru" in email or "@gmail.com" in email or "@yandex.com" in email
        ):
            return True
        else:
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Warning)
            msg.setText("Error: wrong email")
            msg.setWindowTitle("Warning")
            retval = msg.exec_()
            return False

    def check_pathway(self):
        if "pathway" in self.settings.keys():
            return True
        return False

    def load_database(self):
        files_name = os.listdir(self.settings['pathway'])
        self.excel_files_names = os.listdir(self.settings['pathway'])
        for f_index, f in enumerate(files_name, 1):
            excel_file = openpyxl.open(f'{self.settings["pathway"]}/{f}', read_only=True)
            sheet = excel_file.active

            start_line_ind = 0
            while sheet[f'B{start_line_ind}'].value != "№ п/п":
                start_line_ind += 1
            start_line_ind += 3

            end_line_ind = start_line_ind + 1
            while sheet[f'B{end_line_ind}'].value is not None:
                end_line_ind += 1

            equipment = []
            for ind in range(start_line_ind, end_line_ind + 1):
                line = [i.value for i in sheet[f'A{ind}':f'L{ind}'][0]]
                equipment.append(line)

            self.database[str(f_index)] = equipment

    def create_report(self):
        for list_num in self.database.keys():
            equipment_to_report = []
            for eq_num in range(len(self.database[list_num])):
                if self.database[list_num][eq_num][9]:
                    try:
                        if self.database[list_num][eq_num][9].date() < datetime.date.today():
                            equipment_to_report.append(self.database[list_num][eq_num])
                        # добавить колибровку если дата равна и если до даты калибровки осталось меньше двух недель
                    except Exception:
                        print('29 FEBRUARY ERROR')
            self.equipment_report[list_num] = equipment_to_report

        print('-----------------------Список оборудования на калибровку-----------------------')
        pprint.pprint(self.equipment_report)

    def report_to_excel(self):
        return

    def send_mail_with_excel(self, recipient_email, subject, content, excel_file):
        msg = EmailMessage()
        msg['Subject'] = subject
        msg['From'] = self.EMAIL_SENDER
        msg['To'] = recipient_email
        msg.set_content(content)

        with open(excel_file, 'rb') as f:
            file_data = f.read()
        msg.add_attachment(file_data, maintype="application", subtype="xlsx", filename=excel_file)

        with smtplib.SMTP_SSL('smtp.gmail.com', 587) as smtp:
            smtp.login(self.EMAIL_SENDER, self.EMAIL_SENDER_PASS)
            smtp.send_message(msg)

    def send_report(self):
        self.send_mail_with_excel(
            "artjom.verzilov@aquaphor.com", "Test",
            "First Report", "Test.xlsx"
        )
        print("Done")


def excepthook(exc_type, exc_value, exc_tb):
    tb = "".join(traceback.format_exception(exc_type, exc_value, exc_tb))
    print("Oбнаружена ошибка !:", tb)


sys.excepthook = excepthook

if __name__ == '__main__':
    app = QApplication(sys.argv)
    win = MainWindow()
    win.show()
    sys.exit(app.exec())
