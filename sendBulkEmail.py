from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import smtplib
import openpyxl as xl

wb = xl.load_workbook(r'EXCEL_FILE_NAME.xlsx')

myEmailAddress = 'ENTER_YOUR_EMAIL_ADDRESS'
password = 'ENTER_PASSWORD'

'''THIS EXAMPLE USES THE GMAIL SMTP'''
server = smtplib.SMTP('smtp.gmail.com', 587)
server.starttls()
server.login(myEmailAddress, password)

sheet1 = wb.active

emailAddresses = []

'''COLUMN A OF YOUR EXCEL SPREADSHEET SHOULD CONTAIN THE EMAIL ADDRESSES TO WHICH YOU WANT TO SEND EMAILS'''
for cell in sheet1['A']:
    emailAddresses.append(cell.value)

for i in range(len(emailAddresses)):
    msg = MIMEMultipart()
    msg['From'] = myEmailAddress
    msg['To'] = emailAddresses[i]
    msg['Subject'] = '''EMAIL_SUBJECT'''
    text  = '''EMAIL_BODY'''
    msg.attach(MIMEText(text, 'plain'))
    message = msg.as_string()
    server.sendmail(myEmailAddress, emailAddresses[i], message)
    print('Email successfully sent to ' + emailAddresses[i])

server.quit()
