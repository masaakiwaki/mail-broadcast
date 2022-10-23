import pathlib
import csv

WORK_DIR = pathlib.Path()


MAIL_SUBJECT_TEMPLATE_NAME = 'mail_subject_template.txt'
MAIL_SUBJECT_TEMPLATE_PATH =  str(WORK_DIR.cwd() / MAIL_SUBJECT_TEMPLATE_NAME)

MAIL_BODY_TEMPLATE_NAME = 'mail_body_tamplate.txt'
MAIL_BODY_TEMPLATE_PATH =  str(WORK_DIR.cwd() / MAIL_BODY_TEMPLATE_NAME)

MAILING_LIST_TEMPLATE_NAME = 'mailng_list.csv'
MAILING_LIST_TEMPLATE_PATH =  str(WORK_DIR.cwd() / MAILING_LIST_TEMPLATE_NAME)


def import_template():
    with open(MAIL_BODY_TEMPLATE_PATH, encoding="utf-8") as f:
        MAIL_BODY_TEMPLATE= f.read()
    with open(MAIL_SUBJECT_TEMPLATE_PATH, encoding="utf-8") as f:
        MAIL_SUBJECT_TEMPLATE= f.read()
    with open(MAILING_LIST_TEMPLATE_PATH, encoding='shift-jis') as f:
        dict_reader = csv.DictReader(f)
        MAILING_LIST_TEMPLATE = [row for row in dict_reader]
    return MAIL_BODY_TEMPLATE, MAIL_SUBJECT_TEMPLATE, MAILING_LIST_TEMPLATE

