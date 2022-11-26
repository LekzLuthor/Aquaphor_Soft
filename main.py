import sys
import re

from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QApplication, QMainWindow, QTextEdit, QMessageBox
from PyQt5 import uic
import smtplib
from email.message import EmailMessage


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        uic.loadUi('main_window.ui', self)
        self.EMAIL_SENDER = "koshanskiy00@mail.ru"
        self.EMAIL_SENDER_PASS = ""
        self.recipient_email = ""

        self.generateReportButton.clicked.connect(self.send_report)
        self.saveChangesButton.clicked.connect(self.save_changes)
        self.emailSetForm.setPlaceholderText("example@aquaphor.com")

    def save_changes(self):
        email = self.emailSetForm.toPlainText()
        if self.check_email(email):
            self.recipient_email = email
        else:
            return

    @staticmethod
    def check_email(email):
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
