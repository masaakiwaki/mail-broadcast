
import csv
import pathlib
import mail_formatter

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

    mail_to_list = mail_formatter.add_address(mail_to_list , mail_to_add_list )
    mail_cc_list = mail_formatter.add_address(mail_cc_list , mail_cc_add_list )
    mail_bcc_list = mail_formatter.add_address(mail_bcc_list , mail_bcc_add_list )

    mail_cc_list = mail_formatter.deduplication_address(mail_to_list, mail_cc_list)
    mail_bcc_list = mail_formatter.deduplication_address(mail_to_list, mail_bcc_list)
    mail_bcc_list = mail_formatter.deduplication_address(mail_cc_list, mail_bcc_list)

    mail_to = mail_formatter.create_address(mail_to_list)
    mail_cc = mail_formatter.create_address(mail_cc_list)
    mail_bcc = mail_formatter.create_address(mail_bcc_list)

    mail_subject = mail_formatter.create_text(MAIL_SUBJECT_TEMPLATE, case_name, customre_name, mail_body_header)
    mail_body = mail_formatter.create_text(MAIL_BODY_TEMPLATE, case_name, customre_name, mail_body_header)

    mail_formatter.create_mail(mail_to, mail_cc, mail_bcc, mail_subject, mail_body)