import mimetypes
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.base import MIMEBase
from email import encoders
from email.message import EmailMessage
import datetime


def send_to_email():
    server = smtplib.SMTP("smtp.gmail.com", 587)
    server.starttls()
    try:
        server.login('calibration.aquaphor@gmail.com', 'ogwgkvtqnvjsfljr')

        msg = MIMEMultipart()
        msg['Subject'] = f'{datetime.date.today()} Report'
        msg['From'] = 'calibration.aquaphor@gmail.com'

        for file in os.listdir("sours/reports"):
            filename = os.path.basename(file)
            ftype, encoding = mimetypes.guess_type(file)
            file_type, file_subtype = ftype.split('/')

            if file_type == "application":
                with open(f'sours/reports/{file}', encoding="unicode_escape") as f:
                    file = MIMEApplication(f.read(), file_subtype)
            else:
                with open(f'sours/reports/{file}') as f:
                    file = MIMEBase(f.read(), file_subtype)
                    file.set_payload(f.read())
                    encoders.encode_base64(file)

            file.add_header('content-disposition', 'attachment', filename=filename)
            msg.attach(file)

        server.sendmail('calibration.aquaphor@gmail.com', "koshanskiy00@mail.ru", msg.as_string())

        print("The message was sent succesfully")
        return
    except Exception as _ex:
        return f"{_ex}\n"


def send_to_email2():
    SENDER_EMAIL = 'calibration.aquaphor@gmail.com'
    APP_PASSWORD = 'ogwgkvtqnvjsfljr'

    msg = EmailMessage()
    msg['Subject'] = f'{datetime.date.today()} Report'
    msg['From'] = SENDER_EMAIL
    msg['To'] = "koshanskiy00@mail.ru"

    with open("sours/reports/report.xlsx", 'rb') as f:
        file_data = f.read()
    msg.add_attachment(file_data, maintype="application", subtype="xlsx", filename="report.xlsx")

    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login(SENDER_EMAIL, APP_PASSWORD)
        smtp.send_message(msg)


print(send_to_email2())

# if self.reportTime.time().toPyTime() != "00:00:00":
#     self.settings["report_time"] = str(self.reportTime.time().toPyTime())
#     self.timeStatusBar.setText('report time saved')
# elif self.reportTime.time().toPyTime() == self.settings['report_time']:
#     self.timeStatusBar.setText("you haven't entered a new time")
# else:
#     self.timeStatusBar.setText("you haven't entered a new time")
