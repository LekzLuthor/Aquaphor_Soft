import json
import sys
import datetime
import os
import openpyxl
import pprint

from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QApplication, QMainWindow, QTextEdit, QMessageBox, QTimeEdit
from PyQt5 import uic
import smtplib
from email.message import EmailMessage


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        uic.loadUi('main_window.ui', self)
        self.EMAIL_SENDER = "koshanskiy00@mail.ru"
        self.EMAIL_SENDER_PASS = ""
        self.settings = {}
        self.database = {}
        with open("sours/settings.json", "r") as file:  # достаёт настройки из json файла
            self.settings = json.load(file)

        # привязка кнопок + редактура полей
        self.generateReportButton.clicked.connect(self.send_report)
        self.saveChangesButton.clicked.connect(self.save_changes)
        self.emailSetForm.setPlaceholderText("example@aquaphor.com")

        self.test.clicked.connect(self.load_database)

    def save_changes(self):
        email = self.emailSetForm.toPlainText()
        if self.check_email(email):
            self.settings["email"] = email
        else:
            return
        self.settings["report_time"] = str(self.reportTime.time().toPyTime())
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

    def load_database(self):
        files_name = os.listdir("sours/data/")
        for f_index, f in enumerate(files_name):

            print()

            excel_file = openpyxl.open(f'sours/data/{f}', read_only=True)
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

        pprint.pprint(self.database)

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
