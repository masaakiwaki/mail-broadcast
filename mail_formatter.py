import win32com.client


def add_address(address_list, add_list):
    return list(set(address_list + add_list))


def deduplication_address(address_list_1, address_list_2):
    delete_list = list(set(address_list_1) & set(address_list_2))
    for i in delete_list:
        address_list_2.remove(i)
    return address_list_2


def create_address(address_list):
    if len(address_list) == 1:
        return address_list[0]
    else:
        address_name = ''
        for number, i in enumerate(address_list):
            if number == len(address_list) -1:
                address_name += f'{i}'
            else:
                address_name += f'{i}; '
        return address_name
      

def create_text(text_tamplate, case_name, customre_name, mail_body_header):
    i = text_tamplate
    i = i.replace('【case_name】', case_name)
    i = i.replace('【customre_name】', customre_name)
    i = i.replace('【mail_body_header】', mail_body_header)
    return i



def create_mail(mail_to, mail_cc, mail_bcc, mail_subject, mail_body):
    # Outlookのmailオブジェクト設定
    outlook = win32com.client.Dispatch("Outlook.Application")
    objMail = outlook.CreateItem(0) # MailItemオブジェクトのID

    objMail.To = mail_to
    objMail.cc = mail_cc
    objMail.Bcc = mail_bcc
    objMail.Subject = mail_subject
    objMail.Body = mail_body

    objMail.Display(True) # MailItemオブジェクトを画面表示で確認する
    # objMail.Send() # Mailを即時送信
