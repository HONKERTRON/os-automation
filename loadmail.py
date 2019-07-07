import imaplib
import email

def extract_body(payload):
    if isinstance(payload,str):
        return payload
    else:
        return '\n'.join([extract_body(part.get_payload()) for part in payload])

#в body выводится текст сообщения
#в From отправитель
conn = imaplib.IMAP4_SSL("imap.yandex.ru", 993)
conn.login("guap4636@yandex.ru", "4636guap")
conn.select()
typ, data = conn.search(None, 'ALL') #Unseen для непрочитанных
try:
    for num in data[0].split():
        typ, msg_data = conn.fetch(num, '(RFC822)')
        for response_part in msg_data:
            if isinstance(response_part, tuple):
                msg = email.message_from_bytes(response_part[1])
                payload=msg.get_payload()
                body=extract_body(payload)
                print(body)
                From = msg['from']
                print(From)
        #typ, response = conn.store(num, '+FLAGS', r'(\Seen)') пометка сообщения как прочитанного
finally:
    try:
        conn.close()
    except:
        pass
    conn.logout()
