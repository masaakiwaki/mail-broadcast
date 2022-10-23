import pathlib
import csv

WORK_DIR = pathlib.Path()

SELECT_CONFIG_NAME = 'select_template.csv'
SELECT_CONFIG_PATH =  str(WORK_DIR.cwd() / SELECT_CONFIG_NAME)


def create_template_list(SELECT_CONFIG_PATH = SELECT_CONFIG_PATH):
    with open(SELECT_CONFIG_PATH, encoding='shift-jis') as f:
        dict_reader = csv.DictReader(f)
        SELECT_CONFIG_LIST = [row for row in dict_reader]

    template_title_list = []
    for i in SELECT_CONFIG_LIST:
        template_title_list.append(i['title'])
    return  SELECT_CONFIG_LIST, template_title_list,



def import_template(SELECT_CONFIG_LIST, select_config = None):

    for i in SELECT_CONFIG_LIST:
        if select_config:
            if i['title'] == select_config[0]:

                with open(i['config_path'], encoding='shift-jis') as f:
                    dict_reader = csv.DictReader(f)
                    TEMPLATE_LIST = [row for row in dict_reader]

                MAIL_SUBJECT_TEMPLATE_PATH = TEMPLATE_LIST[0]['mail_subject_template'] 
                MAIL_BODY_TEMPLATE_PATH = TEMPLATE_LIST[0]['mail_body_tamplate'] 
                MAILING_LIST_TEMPLATE_PATH = TEMPLATE_LIST[0]['mailng_list'] 

                break

        else:
            MAIL_SUBJECT_TEMPLATE_NAME = 'mail_subject_template.txt'
            MAIL_SUBJECT_TEMPLATE_PATH =  str(WORK_DIR.cwd() / MAIL_SUBJECT_TEMPLATE_NAME)

            MAIL_BODY_TEMPLATE_NAME = 'mail_body_tamplate.txt'
            MAIL_BODY_TEMPLATE_PATH =  str(WORK_DIR.cwd() / MAIL_BODY_TEMPLATE_NAME)

            MAILING_LIST_TEMPLATE_NAME = 'mailng_list.csv'
            MAILING_LIST_TEMPLATE_PATH =  str(WORK_DIR.cwd() / MAILING_LIST_TEMPLATE_NAME)




    with open(MAIL_BODY_TEMPLATE_PATH, encoding="utf-8") as f:
        MAIL_BODY_TEMPLATE= f.read()
    with open(MAIL_SUBJECT_TEMPLATE_PATH, encoding="utf-8") as f:
        MAIL_SUBJECT_TEMPLATE= f.read()
    with open(MAILING_LIST_TEMPLATE_PATH, encoding='shift-jis') as f:
        dict_reader = csv.DictReader(f)
        MAILING_LIST_TEMPLATE = [row for row in dict_reader]
    return MAIL_BODY_TEMPLATE, MAIL_SUBJECT_TEMPLATE, MAILING_LIST_TEMPLATE

