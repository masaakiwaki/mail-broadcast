import mail_formatter
import import_config
import time

import PySimpleGUI as sg


def gui_select_config():

    SELECT_CONFIG_LIST, template_title_list = import_config.create_template_list()
    select_config = ''


    layout = [[sg.Text('Tmplate Select')],
            [sg.Listbox(template_title_list, size=(50, len(template_title_list)), key='-template_name-')],
            [sg.Button('決定'), sg.Button('終了')]]

    window = sg.Window('テンプレート選択', layout)

    while True:
        event, values = window.read()
        if event == sg.WIN_CLOSED or event == '終了':
            break
        if event == '決定':
            select_config = values['-template_name-']
            break

    window.close()

    return SELECT_CONFIG_LIST, select_config










SELECT_CONFIG_LIST, select_config = gui_select_config()

MAIL_BODY_TEMPLATE, MAIL_SUBJECT_TEMPLATE, MAILING_LIST_TEMPLATE = import_config.import_template(SELECT_CONFIG_LIST, select_config)

def make(mail_info):
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



case_name = '案件未入力'
customre_name = 'お客さま未入力'


mail_to_add_list = ['mmm@mmm.com', 'ooo@ooo.com', 'aaa@aaa.com']
mail_cc_add_list = ['aaa@aaa.com']
mail_bcc_add_list = ['yyy@yyy.com']


def add_list(imput_values):
    mail_to_add_list = []
    mail_cc_add_list = []
    mail_bcc_add_list = []
    for i in imput_values:
        if i.startswith('-to'):
            if bool(imput_values[i]):
                mail_to_add_list.append(imput_values[i])
        elif i.startswith('-cc'):
            if bool(imput_values[i]):
                mail_cc_add_list.append(imput_values[i])
        elif i.startswith('-bcc'):
            if bool(imput_values[i]):
                mail_bcc_add_list.append(imput_values[i])
    return mail_to_add_list, mail_cc_add_list,  mail_bcc_add_list



layout = [[sg.Text('案件名', size=(8, 1)), sg.Input(key='-case_name-')],
          [sg.Text('お客さま', size=(8, 1)), sg.Input(key='-customre_name-')],
          [sg.Text('追加アドレス')],
          [],
          [sg.Text('TO', size=(8, 1)), sg.Input(key='-to1-'), sg.Input(key='-to2-'), sg.Input(key='-to3-')],
          [sg.Text('CC', size=(8, 1)), sg.Input(key='-cc1-'), sg.Input(key='-cc2-'), sg.Input(key='-cc3-')],
          [sg.Text('BCC', size=(8, 1)), sg.Input(key='-bcc1-'), sg.Input(key='-bcc2-'), sg.Input(key='-bcc3-')],
          [sg.Text('タイトル')],
          [sg.Input(default_text = MAIL_SUBJECT_TEMPLATE, key='-MAIL_SUBJECT_TEMPLATE-')],
          [sg.Text('本文')],
          [sg.Multiline(default_text = MAIL_BODY_TEMPLATE, key='-MAIL_BODY_TEMPLATE-', size=(100, 20))],
          [sg.Button('決定'), sg.Button('終了')]]

window = sg.Window('メール作成', layout)




while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == '終了':
        break
    if event == '決定':
        case_name = values['-case_name-']
        customre_name = values['-customre_name-']
        MAIL_SUBJECT_TEMPLATE = values['-MAIL_SUBJECT_TEMPLATE-']
        MAIL_BODY_TEMPLATE = values['-MAIL_BODY_TEMPLATE-']

        mail_to_add_list, mail_cc_add_list,  mail_bcc_add_list = add_list(values)
        for mail_info in MAILING_LIST_TEMPLATE:
            make(mail_info)
            time.sleep(3)

window.close()