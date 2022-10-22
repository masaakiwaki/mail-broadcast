import win32com.client
import re

case_name = 'A3カラーレーザープリンター'
customre_name = '株式会社テスト'
mail_body_header = '''株式会社ジャパネットたかた
高田　社長様'''

mail_to = 'xxx@xxx.com; aaa@aaa.com' 
mail_cc = 'yyy@yyy.com'
mail_bcc = 'zzz@zzz.com' 
mail_subject_template = '【お見積依頼】【case_name】　【customre_name】さま分'
mail_body_tamplate = '''【mail_body_header 】

この度は、【case_name】のお見積りをお願いしたくご連絡いたしました。

お客さま: 【customer_name】さま

以上です。

よろしくお願いいたします。'''

#objMail.Attachments.Add('hogehoge.csv') # 送付ファイルがある場合はファイルパスで添付

def create_text(tamplate, case_name, customre_name, mail_body_header):
    i = tamplate.replace('【case_name】', case_name)
    i = i.replace('【customre_name】', customre_name)
    i = i.replace('【mail_body_header】', mail_body_header)
    return i







mail_body = create_text(mail_body_tamplate, case_name, customre_name, mail_body_header)


mail_subject = create_text(mail_subject_template, case_name, customre_name, mail_body_header)

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