import win32com.client
import pathlib

WORK_DIR = pathlib.Path()

MAIL_SUBJECT_TEMPLATE_NAME = 'mail_subject_template.txt'
MAIL_SUBJECT_TEMPLATE_PATH =  str(WORK_DIR.cwd() / MAIL_SUBJECT_TEMPLATE_NAME)
with open(MAIL_SUBJECT_TEMPLATE_PATH, encoding="utf-8") as f:
    MAIL_SUBJECT_TEMPLATE= f.read()

MAIL_BODY_TEMPLATE_NAME = 'mail_body_tamplate.txt'
MAIL_BODY_TEMPLATE_PATH =  str(WORK_DIR.cwd() / MAIL_BODY_TEMPLATE_NAME)
with open(MAIL_BODY_TEMPLATE_PATH, encoding="utf-8") as f:
    MAIL_BODY_TEMPLATE= f.read()

case_name = 'A3カラーレーザープリンター'
customre_name = '株式会社テスト'

mail_body_header = '''株式会社ジャパネットたかた
高田　社長様'''

mail_to_list = ['xxx@xxx.com', 'aaa@aaa.com']
mail_cc_list  = ['yyy@yyy.com']
mail_bcc_list  = ['zzz@zzz.com']

mail_to_add_list = ['mmm@mmm.com', 'ooo@ooo.com', 'aaa@aaa.com']
mail_cc_add_list = ['aaa@aaa.com']
mail_bcc_add_list = ['yyy@yyy.com']

def add_address(address_list, add_list):
    return list(set(address_list + add_list))


mail_to_list = add_address(mail_to_list , mail_to_add_list )
mail_cc_list = add_address(mail_cc_list , mail_cc_add_list )
mail_bcc_list = add_address(mail_bcc_list , mail_bcc_add_list )


def deduplication_address(address_list_1, address_list_2):
    delete_list = list(set(address_list_1) & set(address_list_2))
    for i in delete_list:
        address_list_2.remove(i)
    return address_list_2


mail_cc_list = deduplication_address(mail_to_list, mail_cc_list)
mail_bcc_list = deduplication_address(mail_to_list, mail_bcc_list)
mail_bcc_list = deduplication_address(mail_cc_list, mail_bcc_list)


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
      


mail_to = create_address(mail_to_list)
mail_cc = create_address(mail_cc_list)
mail_bcc = create_address(mail_bcc_list)



def create_text(text_tamplate, case_name, customre_name, mail_body_header):
    i = text_tamplate
    i = i.replace('【case_name】', case_name)
    i = i.replace('【customre_name】', customre_name)
    i = i.replace('【mail_body_header】', mail_body_header)
    return i

mail_subject = create_text(MAIL_SUBJECT_TEMPLATE, case_name, customre_name, mail_body_header)
mail_body = create_text(MAIL_BODY_TEMPLATE, case_name, customre_name, mail_body_header)

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

create_mail(mail_to, mail_cc, mail_bcc, mail_subject, mail_body)