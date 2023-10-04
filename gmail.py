from random import randint
import pandas as pd
from email.message import EmailMessage
import smtplib
import logging
import time
import sys
import pdfkit
import imgkit
from operator import index
from random import randint, choices
from email import encoders
import pandas as pd
from email.message import EmailMessage
import smtplib
import logging
import time
import sys
import os
import string
paths=os.getcwd()
os.chdir(paths)


#path_wkhtmltopdf = r'C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe'
#path_wkhtmltopdf = '/usr/local/bin/wkhtmltopdf'
#config = pdfkit.configuration(wkhtmltopdf=path_wkhtmltopdf)

logging.basicConfig(filename='mail.log', level=logging.DEBUG)

totalSend = 1
if(len(sys.argv) > 1):
    totalSend = int(sys.argv[1])

email_df = pd.read_csv('mail.csv')
contactsData = pd.read_csv('contacts.csv')
subjects = pd.read_csv("subjects.csv")
body_data = pd.read_csv("body.csv")
html_data = pd.read_csv("invoice.csv")
from_name = pd.read_csv("fromNames.csv")


def send_mail(fname, email, emailId, password, body_data, subjectPhrase, fromName, html_data):
    print(email, html_data)
    newMessage = EmailMessage()

    # Invoice Number and Subject
    invoiceNo = randint(1000000000000, 9999999999999)
    invoiceNoid = randint(100000000, 999999999)
    randomString = ''.join(choices(string.ascii_uppercase, k=4))
    randomString2 = ''.join(choices(string.ascii_uppercase, k=6))
    randomString3 = ''.join(choices(string.ascii_uppercase, k=4))
    #subject = subjectPhrase+ " "+ "P4NNOJRL0BO74F5 "
    subject = subjectPhrase+ " " +randomString2+randomString3+str(invoiceNo) 
    num = randint(1111, 9999)
    numm = randint(1111111111, 9999999999)
    newMessage['Subject'] = subject
    newMessage['From'] = f"{fromName}#{randomString3}{num}<{emailId}>"
    newMessage['To'] = email
    transaction_id = randint(10000000, 99999999)

    # Mail Body Content
    body = open('body.txt', 'r').read()
    body = body.replace('$fname', fname)
    body = body.replace('$email', email)
    body = body.replace('$invoice_no', str(transaction_id))
    body = body.replace('$random', str(randomString))
    body = body.replace('$ran',  str(randomString3))
    body = body.replace('$receipt',  str(numm))
    body = body.replace('$body', body_data['body'])
    body = body.replace('$note', body_data['note'])

    # Mail PDF File\
    html = open('html_code.html', 'r').read()
    original_html = html
    html = html.replace('$logo', html_data['logo'])
    html = html.replace('$name', html_data['name'])
    html = html.replace('$fname', fname)
    html = html.replace('$email', email) 
    html = html.replace('$lastname', html_data['lastname'])
    html = html.replace('$body', html_data['body'])
    html = html.replace('$body', body_data['body'])
    html = html.replace('$product', html_data['product'])
    html = html.replace('$quantity', str(html_data['quantity']))
    html = html.replace('$price', html_data['price'])
    html = html.replace('$invoice_no', str(transaction_id))
    html = html.replace('$random',  str(randomString))
    html = html.replace('$receipt',  str(numm))
    html = html.replace('$ran',  str(randomString3))
    

    # saving the changes to html_code.html
    with open('html_code.html', 'w') as f:
        f.write(html)

    file = str(num) +randomString  + ".pdf"
    pdfkit.from_file('html_code.html', file)

    html = open('html_code.html', 'r').read()
    html = html.replace(str(transaction_id), '$invoice_no')
    # reverting change to original html content to html_code.html
    with open('html_code.html', 'w') as f:
        f.write(original_html)

    newMessage.set_content(body)

    try:
        with open(file, 'rb') as f:
            file_data = f.read()
            file_name = f.name
        newMessage.add_attachment(
            file_data, maintype='application', subtype='octet-stream', filename=file_name)
            
        with smtplib.SMTP('smtp.gmail.com', 587) as smtp:
            smtp.starttls()
            smtp.login(emailId, password)
            smtp.send_message(newMessage)
            smtp.quit()
            
        print("removing file: ", file, file_name)
        os.remove(file)
        print("file removed")
        print(f"send to {email} by {emailId} successfully : {totalSend}")
        logging.info(
            f"send to {email} by {emailId} successfully : {totalSend}")

    except smtplib.SMTPResponseException as e:
        error_code = e.smtp_code
        error_message = e.smtp_error
        print(f"send to {email} by {emailId} failed")
        logging.info(f"send to {email}  by {emailId} failed")
        print(f"error code: {error_code}")
        print(f"error message: {error_message}")
        logging.info(f"error code: {error_code}")
        logging.info(f"error message: {error_message}")

        remove_email(emailId, password)


def start_mail_system():
    global totalSend
    j = 0  # for sender emails
    k = 0  # for bodies
    l = 0  # for subjects
    m = 0  # for From Name
    n = 0  # for invoice data

    for i in range(len(contactsData)):
        email_df = pd.read_csv('mail.csv')
        if(j >= len(email_df)):
            j = 0
        time.sleep(0.6)
        try:
            send_mail(contactsData.iloc[i]['fname'], contactsData.iloc[i]['email'], email_df.iloc[j]['email'],
                          email_df.iloc[j]['password'], body_data.iloc[k], subjects.iloc[l]['subject'], from_name.iloc[m]['fromName'], html_data.iloc[n])
            totalSend += 1
        except Exception as e:
            print(e)
            print(f"Not able to send mail to {contactsData.iloc[i]['fname']}, {contactsData.iloc[i]['email']}")
        j = j + 1
        k = k + 1
        l = l + 1
        m = m + 1
        n = n + 1

        if j == len(email_df):
            j = 0
        if k == len(body_data):
            k = 0
        if l == len(subjects):
            l = 0
        if m == len(from_name):
            m = 0
        if n == len(html_data):
            n = 0
    quit()


def remove_email(emailId, password):
    df = pd.read_csv('mail.csv')
    index = df[df['email'] == emailId].index
    df.drop(index, inplace=True)
    df.to_csv('mail.csv', index=False)
    print(f"{emailId} removed from mail.csv")
    logging.info(f"{emailId} removed from mail.csv")


try:
    print("Beta Mailer")
    for i in range(6):
        start_mail_system()
except KeyboardInterrupt as e:
    print(f"\n\ncode stopped by user")
