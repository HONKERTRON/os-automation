import imaplib
import email

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
    list_text = []
    list_from = []
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
                    list_text.append(Text)
                    list_from.append(From)
            typ, response = conn.store(num, '+FLAGS', r'(\Seen)') #пометка сообщения как прочитанного
    finally:
        try:
            conn.close()
        except:
            pass
        conn.logout()
        list_mail.append(list_from)
        list_mail.append(list_text)
        return list_mail

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
