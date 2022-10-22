import win32com.client
import csv
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

mail_to_add_list = ['mmm@mmm.com', 'ooo@ooo.com', 'aaa@aaa.com']
mail_cc_add_list = ['aaa@aaa.com']
mail_bcc_add_list = ['yyy@yyy.com']

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



with open('mailng_list.csv', encoding='shift-jis') as f:
    dict_reader = csv.DictReader(f)
    mailing_list = [row for row in dict_reader]


for mail_info in mailing_list:
    
    mail_body_header = ''
    mail_to_list = []
    mail_cc_list  = []
    mail_bcc_list  = []

    for i in mail_info :
        if i == 'header':
            mail_body_header = mail_info[i]
        elif i.startswith('to'):
            if bool(mail_info[i]):
                mail_to_list.append(mail_info[i])
        elif i.startswith('cc'):
            if bool(mail_info[i]):
                mail_cc_list.append(mail_info[i])
        elif i.startswith('bcc'):
            if bool(mail_info[i]):
                mail_bcc_list.append(mail_info[i])

    mail_to_list = add_address(mail_to_list , mail_to_add_list )
    mail_cc_list = add_address(mail_cc_list , mail_cc_add_list )
    mail_bcc_list = add_address(mail_bcc_list , mail_bcc_add_list )

    mail_cc_list = deduplication_address(mail_to_list, mail_cc_list)
    mail_bcc_list = deduplication_address(mail_to_list, mail_bcc_list)
    mail_bcc_list = deduplication_address(mail_cc_list, mail_bcc_list)

    mail_to = create_address(mail_to_list)
    mail_cc = create_address(mail_cc_list)
    mail_bcc = create_address(mail_bcc_list)

    mail_subject = create_text(MAIL_SUBJECT_TEMPLATE, case_name, customre_name, mail_body_header)
    mail_body = create_text(MAIL_BODY_TEMPLATE, case_name, customre_name, mail_body_header)

    create_mail(mail_to, mail_cc, mail_bcc, mail_subject, mail_body)