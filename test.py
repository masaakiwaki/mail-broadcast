import win32com.client

mail_to = 'xxx@xxx.com; aaa@aaa.com' 
mail_cc = 'yyy@yyy.com'
mail_bcc = 'zzz@zzz.com' 
mail_subject = 'title'
mail_body = 'mail body test + sign'
#objMail.Attachments.Add('hogehoge.csv') # 送付ファイルがある場合はファイルパスで添付


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