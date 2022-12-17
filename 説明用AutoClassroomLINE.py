#モジュールをインポート
import imapclient
from backports import ssl
from OpenSSL import SSL
import pyzmail
import time
import pyautogui
import pandas as pd
import keyboard
import requests

#メールをチェックする関数を定義
def check_mail():
    global imap
    
    title = ""

    #解析メールの結果保存用
    From_list = []
    Cc_list = []
    Bcc_list = []
    Subject_list = []
    Body_list = []
    context = ssl.SSLContext(SSL.TLSv1_2_METHOD)

    imap = imapclient.IMAPClient("imap.gmail.com",ssl=True,ssl_context=context)

    my_mail = "ここにメールアドレスを入力" #メールアドレス
    app_password = "ここにアプリパスワードを入力（2段階認証をし、gmailようのアプリパスワードを作成）"    #パスワード

    imap.login(my_mail,app_password)    #imapにログインする。

    imap.select_folder("INBOX",readonly=False)   #これがないと、動かない。

    KWD = imap.search(["UNSEEN","FROM","no-reply@classroom.google.com"])    #検索のキーワード。
    #"UNSEEN" 未読/"FROM","no-reply@classroom.google.com" クラスルームの通知を送ってくるメールアドレスを指定
    
    raw_message = imap.fetch(KWD,["BODY[]"])

    #検索結果保存
    for j in range(len(KWD)):
        
       #特定メール取得
        message = pyzmail.PyzMessage.factory(raw_message[KWD[j]][b"BODY[]"])
        
       #件名取得
        Subject = message.get_subject()
        Subject_list.append(Subject)
        #本文
        Body = message.text_part.get_payload().decode(message.text_part.charset)
        Body_list.append(Body)
    
    #titleにメールの本文を入れる。
    for title in Body_list:
        print("メールを受け取りました。")   #確認のため、「メールを受け取りました。」と表示

    return title    #この関数を実行したときの戻り値をtitleとする。

title = ""
TOKEN = "LINE notifyのトークンを入力"   #トークン
api_url = "https://notify-api.line.me/api/notify"
count = 0

print("自動でGoogleClassroomの投稿をLINEに投稿するプログラムです。")    #コードを実行したとき、最初に表示される。
print("Enterキーを押すと終了します。")  #説明

while True: #無限ループ
    title = check_mail()    #titleに関数check_mail()の戻り値を代入
    if title == "": #titleの値が何もなかったら
        title = ""
    else:   #titleに何か値が入っていたら
        title = title[title.find("新しいお知らせ"):]    #本文をトリミング
        title = title[:title.find("開く")]  #本文をトリミング
        print(title)    #表示する。
        #LINEで送信するプログラム
        send_contents = title   #通知内容
        TOKEN_dic = {"Authorization": "Bearer" + " " + TOKEN}
        send_dic = {"message": send_contents}
        requests.post(api_url, headers=TOKEN_dic, data=send_dic)
        
        count = count + 1   #リザルトとして表示するため
        title = ""        
    if keyboard.is_pressed("enter"):    #enterキーが押されたら、終了する
        print('メッセージ','処理を抜けました。')
        break

print("メッセージ送信回数 : ", count)   #リザルトを表示

#print(title)
#titleは本文