import zipfile
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import ntpath
from datetime import datetime
import os, shutil, sys


def get_path_of_the_local_directory():
    return os.path.abspath(os.path.dirname(sys.argv[0]))


def zipdir(path, ziph):
    # ziph is zipfile handle
    for root, dirs, files in os.walk(path):
        for file in files:
            ziph.write(os.path.join(root, file))


def pack_to_zip(xlsx_name, foldercalls_name, with_xlsx=True, with_folder=True):
    filename = xlsx_name+'.zip'

    jungle_zip = zipfile.ZipFile(filename, 'w', zipfile.ZIP_DEFLATED)
    if with_xlsx:
        jungle_zip.write(xlsx_name+'.xlsx')
    if with_folder:
        zipdir(foldercalls_name+'\\', jungle_zip)
    jungle_zip.close()

    return filename


def send_email(_email, filename):
    fromaddr = "atsserverdotnet@gmail.com"
    toaddr = _email

    # instance of MIMEMultipart
    msg = MIMEMultipart()

    # storing the senders email address
    msg['From'] = fromaddr

    # storing the receivers email address
    msg['To'] = toaddr

    # storing the subject
    msg['Subject'] = "Отчёт по звонкам АТС за прошедший месяц"

    # string to store the body of the mail
    body = ""

    # attach the body with the msg instance
    msg.attach(MIMEText(body, 'plain'))

    # open the file to be sent
    attachment = open(get_path_of_the_local_directory()+'\\'+filename, "rb")

    # instance of MIMEBase and named as p
    p = MIMEBase('application', 'octet-stream')

    # To change the payload into encoded form
    p.set_payload(attachment.read())

    # encode into base64
    encoders.encode_base64(p)

    p.add_header('Content-Disposition', "attachment; filename={}".format('Otchet.xlsx'))

    # attach the instance 'p' to instance 'msg'
    msg.attach(p)

    # sending the mail
    server = smtplib.SMTP_SSL('smtp.gmail.com:465')
    server.login('******************')', '******************')
    server.sendmail('******************')', _email, msg.as_string())

    # terminating the session
    server.quit()


def delete_files(xlsxname, foldercalls):
    shutil.rmtree(foldercalls)
    os.remove(xlsxname+'.xlsx')

#def
