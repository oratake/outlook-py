import win32com.client
import mail_secret

outlook = win32com.client.Dispatch("Outlook.Application")
mail = outlook.CreateItem(0)  # 0でメールオブジェクトの生成

# メール生成
mail.to = mail_secret.mail_to
mail.cc = mail_secret.mail_cc

inputDay = input('日付を入力: ')
mail.subject = inputDay + '管理部配送予定'
mail.bodyFormat = 1  # 1:テキストとして送信
mail.body = mail_secret.mail_body(inputDay)

mail.display(True)