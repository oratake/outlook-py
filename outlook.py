import pdb  # デバッグ用
# pdb.set_trace()
import win32com.client
import mail_secret
from tkinter import filedialog, Tk
import sys
import datetime
import re  # 正規表現
import os
import shutil
from datetime import datetime

outlook = win32com.client.Dispatch("Outlook.Application")
mail = outlook.CreateItem(0)  # 0でメールオブジェクトの生成

# メール生成
mail.to = mail_secret.mail_to
mail.cc = mail_secret.mail_cc

# 日付を取得しバリデーション
input_day = input('日付を入力(例 12/3): ')
DAY_PATTERN = '\d{1,2}/\d{1,2}'
repatter = re.compile(DAY_PATTERN)
result = repatter.match(input_day)

if not(result):
    print('[ERROR] 例のように日付を入力してください')
    exit(0)

date_obj = datetime.strptime(input_day, '%m/%d')

print(result.group())

# 添付するPDFのファイルパスを取得
root = Tk()
root.withdraw()
src_file = filedialog.askopenfilename(filetypes=[("PDFファイル", "*.pdf")], title="スキャンしたpdfファイルを選択")

if src_file == "":
    print("ファイルが選択されませんでした。")
    sys.exit()

# PDFファイルを実行中のpythonファイルと同じディレクトリにコピー
dst_file = os.path.join(os.path.dirname(__file__), "管理部配送予定" + date_obj.strftime('%m%d') + '.pdf')
shutil.copy2(src_file, dst_file)

mail.Attachments.Add(dst_file)

mail.subject = input_day + '管理部配送予定'
mail.bodyFormat = 1  # 1:テキストとして送信
mail.body = mail_secret.mail_body(input_day)

mail.display(False)