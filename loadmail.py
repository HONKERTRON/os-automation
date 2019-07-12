import imaplib
import email
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from oauth2client.service_account import ServiceAccountCredentials
import gspread
import settings
import time


def extract_body(payload):
    if isinstance(payload,str):
        return payload
    else:
        return '\n'.join([extract_body(part.get_payload()) for part in payload])

def get_first_text_block(self, email_message_instance):
    maintype = email_message_instance.get_content_maintype()
    if maintype == 'multipart':
        for part in email_message_instance.get_payload():
            if part.get_content_maintype() == 'text':
                return part.get_payload()
    elif maintype == 'text':
        return email_message_instance.get_payload()

def get_list_email(self):
    list_mail = []
    mail = []
    conn = imaplib.IMAP4_SSL("imap.yandex.ru", 993)
    conn.login("guap4636@yandex.ru", "4636guap")
    conn.select()
    typ, data = conn.search(None, 'Unseen')  # Unseen для непрочитанных
    try:
        for num in data[0].split():
            typ, msg_data = conn.fetch(num, '(RFC822)')
            for response_part in msg_data:
                if isinstance(response_part, tuple):
                    msg = email.message_from_bytes(response_part[1])
                    payload = msg.get_payload()
                    body = extract_body(payload)
                    From = msg['from']
                    Text = get_first_text_block(msg)
                    mail.append(Text)
                    mail.append(From)
            typ, response = conn.store(num, '+FLAGS', r'(\Seen)') #пометка сообщения как прочитанного
    finally:
        try:
            conn.close()
        except:
            pass
        conn.logout()
        list_mail.append(mail)
        return list_mail


def send_email(to_addr, body_text):
    msg = MIMEMultipart()
    msg['From'] = "guap4636@yandex.ru"
    msg['To'] = to_addr
    msg['Subject'] = "Операционные системы"
    msg.attach(MIMEText(body_text, 'plain'))

    smtpObj = smtplib.SMTP_SSL('smtp.yandex.ru:465')
    smtpObj.login('guap4636@yandex.ru', '4636guap')
    text = msg.as_string()
    smtpObj.sendmail("guap4636@yandex.ru", to_addr, text)
    smtpObj.quit()

#возвращает корректные данные из письма в виде для поиска mail_data[0] - вариант
#mail_data[1] - группа mail_data[2] - ФИО все строчными буквами mail_data[3] - ник на GitHub
def mail_processing (mail_str):
    mail_str = mail_str.strip()
    mail_str = mail_str.lower()
    mail_content = mail_str.split(' ')
    fio = mail_content[2] + " " + mail_content[3] + " " + mail_content[4]
    mail_data = []
    mail_data.append(mail_content[0])
    mail_data.append(mail_content[1])
    mail_data.append(fio)
    mail_data.append(mail_content[5])
    return mail_data

def add_to_gsheets(mail_list):
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name(settings.gsheet_key_filename, scope)
    conn = gspread.authorize(creds)
    for mail in mail_list:
        pr_mail = mail_processing(mail[0])
        num_var = pr_mail[0]
        group_name = pr_mail[1]
        stud_fio = pr_mail[2]
        git_name = pr_mail[3]
        try :
            worksheet = conn.open(settings.gspreadsheet_name).worksheet(group_name)
            time.sleep(1)
        except :
            send_email(mail[1], "Неверно указана группа. Проверьте данные и отправьте повторно.")
            raise Exception("No group {}: {}".format(group_name, mail))
        names_list = [x.lower() for x in worksheet.col_values(2)[2:]]
        time.sleep(1)
        if stud_fio in names_list:
            stud_row = names_list.index(stud_fio) + 3
        else :
            send_email(mail[1], "Неверно указано ФИО. Проверьте данные и отправьте повторно.")
            raise Exception("No student {}: {}".format(stud_fio, mail))
        if num_var != worksheet.cell(stud_row, 1).value.strip() :
            send_email(mail[1], "Внимание ваш вариант "+worksheet.cell(stud_row, 1).value.strip())
        is_empty = worksheet.cell(stud_row, 19).value.strip() == ''
        if is_empty:
            worksheet.update_cell(stud_row, 19, git_name)

#в body выводится текст сообщения
#в From отправитель
#conn = imaplib.IMAP4_SSL("imap.yandex.ru", 993)
#conn.login("guap4636@yandex.ru", "4636guap")
#conn.select()
#typ, data = conn.search(None, 'ALL') #Unseen для непрочитанных
#try:
#    for num in data[0].split():
#        typ, msg_data = conn.fetch(num, '(RFC822)')
#        for response_part in msg_data:
#            if isinstance(response_part, tuple):
#                msg = email.message_from_bytes(response_part[1])
#                payload=msg.get_payload()
#                body=extract_body(payload)
#                print(body)
#                From = msg['from']
#                print(From)
#                Text = get_first_text_block(msg)
#                print(Text)
#        #typ, response = conn.store(num, '+FLAGS', r'(\Seen)') пометка сообщения как прочитанного
#finally:
#    try:
#        conn.close()
#    except:
#        pass
#    conn.logout()

#addr = "alexsoldat98@yandex.ru"
#body = "test message"
#send_email(addr, body)