#!/usr/bin/env python
# coding: utf-8


import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText#專門傳送正文
from email.mime.image import MIMEImage
import smtplib
import datetime


#讀excel
def load_data(file_name):
    df_orderData = pd.read_excel(file_name)
    df_data = pd.DataFrame(df_orderData,columns=['OEB01','OEB03','OEB15','OEB16','OEA02','OEA14'])
    df_data["pk"] = df_data["OEB01"] + "_" + df_data["OEB03"].map(str)
    return df_data

# 資料處理
def data_process(df, df1):
    tmp = df.append(df1)
    tmp.drop_duplicates(keep=False, inplace=True)
    df2=pd.DataFrame(tmp)

    last = df2.drop_duplicates(subset=['pk'],keep='last')
    first = df2.drop_duplicates(subset=['pk'],keep='first')
    dupl = df2.drop_duplicates(subset=['pk'],keep=False)

    new = last.append(dupl).drop_duplicates(subset=['pk'],keep=False)
    old = first.append(dupl).drop_duplicates(subset=['pk'],keep=False)

    final = new.merge(old,on = ['pk','OEB01','OEB03','OEB15','OEA02','OEA14'], suffixes=('_new','_old') )
    final = final[['OEA14','OEA02','OEB01','OEB03','OEB15','OEB16_old','OEB16_new']]
    final.sort_values('OEA14')
    final.rename(columns={'OEA02': '訂單日期', 'OEA14': '業務編號', 'OEB01': '訂單號碼', 'OEB03': '項次', 'OEB15': '約定交貨日', 'OEB16_old': '原排定交貨日', 'OEB16_new': '新排定交貨日'}, inplace=True)

    #dataframe轉html
    df_html = final.to_html(escape=False,index=False, justify = 'center')

    if old.equals(new):
        return 0, df_html
    else:
        return 1, df_html

def read_file(file):
    with open(file,'r') as f:
            content_list = f.read()
    return content_list

def SendMail(df_html):
    #print('start to send mail...')
    sender = "hl94vul3h6@gmail.com" #step1:setup sender gmail,ex:"Fene1977@superrito.com"
    password = "eeecnzvlhavhksyb" #step2:setup sender gmail password
    recipients = read_file('收件人.txt') #step3:setup recipients mail
    today_date = datetime.date.today()
    sub = today_date.strftime("%m/%d") + "訂單交期更改通知" #step4:setup your subject
    
    #單個收件人,ex:recipients = 'Fene1977@dayrep.com'  
    #多個收件人,ex:"Hinte1969@jourrapide.com,Fene1977@rhyta.com,Fene1977@teleworm.us"
    
    outer = MIMEMultipart()
    outer['From'] = sender #step:setup sender gmail
    outer['To'] = recipients #step:setup recipient mail
    #outer["Cc"]="cc mail" #step:setup cc mail
    outer['Subject'] = sub #step:setup your subject

    #設定純文字資訊
    plainText = "偵測到訂單排定交貨日更改，請確認以下訂單："
    msgText = MIMEText(plainText, 'plain', 'utf-8')
    outer.attach(msgText)
    #設定HTML資訊
    htmlText = df_html #step7:edit your mail content
    msgText = MIMEText(htmlText, 'html', 'utf-8')
    outer.attach(msgText)

    mailBody = outer.as_string()
    #-----------------------------------------------------------------------
    
    # 寄送EMAIL
    try:
        with smtplib.SMTP(host="smtp.gmail.com", port="587") as s: #send webservice to gmail smtp socket
            s.ehlo()  # 驗證SMTP伺服器
            s.starttls()  # 建立加密傳輸
            s.login(sender, password)  # 登入寄件者gmail
            s.sendmail(sender, recipients,mailBody)  # 寄送郵件
            s.close()
        print("Email sent!")
    except:
        print("Unable to send the email. Error: ", sys.exc_info()[0])
        raise
        
def main():
    yesterday = load_data('訂單test.xlsx')
    today = load_data('訂單test1.xlsx')
    flag, mail_content = data_process(yesterday, today)

    if flag == 0:
        print("未更改")
    else:
        SendMail(mail_content)
        print("交期已被更改")

main()
