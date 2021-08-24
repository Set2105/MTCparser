import imaplib
import email
import re


def imap_login(mail, password, server_address):
    print("Авторизация {}\n Логин:{}\n Пароль:{}".format(server_address, mail, password))
    try:
        mail_obj = imaplib.IMAP4_SSL(server_address, 993)
        mail_obj.login(mail, password)
        print("Авторизация успешна.")
        return mail_obj
    except:
        print("Авторизация неуспешна!")
        return None


def find_mts_message(subject_text, mail_obj):
    if mail_obj:
        mail_obj.select("inbox")
        result, data = mail_obj.search(None, "ALL")
        id_list = data[0].split()
        for i in reversed(range(len(id_list))):
            post_id = i
            result, data = mail_obj.fetch(id_list[post_id], "(RFC822)")  # Получаем тело письма (RFC822) для данного ID
            # Тело письма в необработанном вид, включает в себя заголовки и альтернативные полезные нагрузки
            raw_email = data[0][1]
            email_message = email.message_from_bytes(raw_email)
            sub = email_message['Subject']
            if sub[0] == '=' and sub[1] == '?':
                sub = email.header.decode_header(sub)
                sub = sub[0][0].decode(sub[0][1])
            if sub == subject_text:
                pure_html_text_message = get_first_text_block(email_message)
                print("Найдено письмо с темой \"{}\".".format(subject_text))
                return pure_html_text_message
    else:
        print("Объект mail_obj не получен !")
        return None
    print("В обозначенном диапазоне письмо с темой \"{}\" не найдено !".format(subject_text))
    return None


def get_first_text_block(email_message_instance):
    try:
        maintype = email_message_instance.get_content_maintype()
        if maintype == 'multipart':
            for part in email_message_instance.get_payload():
                if part.get_content_maintype() == 'text':
                    return part.get_payload(decode=True)
        elif maintype == 'text':
            return email_message_instance.get_payload(decode=True)
    except:
        print("В письме не найден текстовый блок !")
        return None


def formatting_post_text(text):
    address = re.search('адресу:  [А-я \-,0-9(:]+\)', text)[0]
    address = re.split('  ', address)[1]
    sale_code = re.search('код продажи:[0-9]+', text)[0]
    sale_code = re.split(':', sale_code)[1]
    return sale_code


def error_result(order_num_list):
    buy_num = {}
    result = {}
    for key in order_num_list:
        buy_nums = {}
        for value in order_num_list[key]:
            buy_nums.update({value: '-Error-'})
        result.update({key: buy_nums})
    return result


# {idТелефона: {ордер1: байкод1, ордер2: байкод2}, ...}
def imap_script_main(login, password, imap_srv, order_num_list):
    result = {}
    print(login, password, imap_srv)
    try:
        mail = imap_login(login, password, imap_srv)
        for key in order_num_list:
            buy_nums = {}
            for value in order_num_list[key]:
                order_num = value
                if mail:
                    print(order_num)
                    post_text = find_mts_message("Ваш заказ №{} готов к выдаче!".format(order_num), mail)
                    if post_text:
                        post_text = post_text.decode('utf-8')
                        buy_num = formatting_post_text(post_text)
                        buy_nums.update({value: buy_num})
                    else:
                        buy_nums.update({value: 'NotFound'})
                result.update({key: buy_nums})
        mail.close()
        mail.logout()
    except:
        result = error_result(order_num_list)
    if not result:
        result = error_result(order_num_list)
    print(result)
    return result




