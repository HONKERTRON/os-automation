import imaplib
import email
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from bs4 import BeautifulSoup
from oauth2client.service_account import ServiceAccountCredentials
import gspread
import settings
import time


#порерка на тип содержимого и извлечение текста из тела письма
def get_first_text_block(email_message_instance):
    maintype = email_message_instance.get_content_maintype()
    if maintype == 'multipart':
        for part in email_message_instance.get_payload():
            if part.get_content_maintype() == 'text':
                return part.get_payload()
    elif maintype == 'text':
        return email_message_instance.get_payload()

#получает лист писем с почты из непрочитанных
#list_mail[0] - тело письма list_mail[1] - от кого письмо
def get_list_email():
    list_mail = []
    mail = []
    conn = imaplib.IMAP4_SSL("imap.yandex.ru", 993)
    conn.login("alexsoldat98@yandex.ru", "SamickLK-35A")
    conn.select()
    typ, data = conn.search(None, 'Unseen')  # Unseen для непрочитанных
    try:
        for num in data[0].split():
            mail = []
            typ, msg_data = conn.fetch(num, '(RFC822)')
            for response_part in msg_data:
                if isinstance(response_part, tuple):
                    msg = email.message_from_bytes(response_part[1])
                    payload = msg.get_payload(decode=True)
                    From = msg['from']
                    Text = get_first_text_block(msg)
                    buf = BeautifulSoup(Text, "lxml").text
                    mail.append(buf)
                    mail.append(From)
            list_mail.append(mail)
            typ, response = conn.store(num, '+FLAGS', r'(\Seen)') #пометка сообщения как прочитанного
    finally:
        try:
            conn.close()
        except:
            pass
        conn.logout()
        return list_mail

#отправление письма
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
    smtpObj.sendmail("guap4636@yandex.ru", "guap4636@yandex.ru", text) #отправка почты себе, тк не сохранятюся в отправленных
    smtpObj.quit()

#возвращает корректные данные из письма в виде для поиска mail_data[0] - группа mail_data[1] - вариант
# mail_data[2] - ФИО все строчными буквами mail_data[3] - ник на GitHub
def mail_processing (mail_str):
    mail_str = mail_str.strip()
    mail_str = mail_str.lower()
    if mail_str.count(' ') >= 4:
        mail_content = mail_str.split(' ')
        fio = mail_content[2] + " " + mail_content[3] + " " + mail_content[4]
        mail_data = []
        mail_data.append(mail_content[0].strip())
        mail_data.append(mail_content[1].strip())
        mail_data.append(fio.strip())
        mail_data.append(mail_content[5].strip())
    elif mail_str.count('\n') >= 4:
        mail_content = mail_str.split('\n')
        fio = mail_content[2] + " " + mail_content[3] + " " + mail_content[4]
        mail_data = []
        mail_data.append(mail_content[0].strip())
        mail_data.append(mail_content[1].strip())
        mail_data.append(fio.strip())
        mail_data.append(mail_content[5].strip())
    return mail_data

#добавляет ник гита в гугл док
def add_to_gsheets(mail_list):
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name(settings.gsheet_key_filename, scope)
    conn = gspread.authorize(creds)
    git_niks = []
    sheets = conn.open(settings.gspreadsheet_name).worksheets()
    for sh in sheets:
        git_niks.append(sh.col_values(19)[2:])
    for mail in mail_list:
        pr_mail = mail_processing(mail[0])
        num_var = pr_mail[1]
        group_name = pr_mail[0]
        stud_fio = pr_mail[2]
        git_name = pr_mail[3]

        try :
            worksheet = conn.open(settings.gspreadsheet_name).worksheet(group_name)
            time.sleep(1)
        except :
            send_email(mail[1], "Неверно указана группа. Проверьте данные и отправьте повторно.")
            break
        names_list = [x.lower() for x in worksheet.col_values(2)[2:]]
        time.sleep(1)
        if stud_fio in names_list:
            stud_row = names_list.index(stud_fio) + 3
        else :
            send_email(mail[1], "Неверно указано ФИО. Проверьте данные и отправьте повторно.")
            break
        if num_var != worksheet.cell(stud_row, 1).value.strip() :
            send_email(mail[1], "Внимание ваш вариант "+worksheet.cell(stud_row, 1).value.strip())

        not_free_nick = git_name in git_niks
        if not_free_nick:
            send_email(mail[1], "Этот профиль GitHub занят!")
            break
        is_empty = worksheet.cell(stud_row, 19).value.strip() == ''
        if is_empty and not_free_nick != True:
            worksheet.update_cell(stud_row, 19, git_name)
            time.sleep(1)

