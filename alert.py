#!/usr/bin/env python
# coding: utf-8

import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText #專門傳送正文
import smtplib
import sys
import datetime
import os
import shutil
import warnings
from tqdm import tqdm

#讀excel
def load_data(file_name, sheet_num=0):
    with warnings.catch_warnings(record=True):  #消除warning
        warnings.simplefilter("always")
        try:
            for i in tqdm(range(0,1), ncols = 100, desc = f"Loading {file_name} ..."):
                df_orderData = pd.read_excel(file_name, sheet_name = sheet_num, engine="openpyxl")
            df_data = pd.DataFrame(df_orderData,columns=['OEB01','OEB03','OEB15','OEB16','OEA02','OEA14','TC_SFA104'])
            df_data["pk"] = df_data["OEB01"] + "_" + df_data["OEB03"].map(str) # pk(不能重複) = 訂單編號 + 項次
            return df_data
        except Exception as e:
            print(e)


# 資料處理
def data_process(df, df1):
    # 可能被更改的欄位: OEB16、TC_SFA104
    # 找昨天跟今天所有不重複的資料(含更改前後pk相同的訂單)
    tmp = df.append(df1)
    tmp.drop_duplicates(keep=False, inplace=True)
    all_diff =pd.DataFrame(tmp) 

    # 找昨天跟今天所有pk不重複的資料
    last = all_diff.drop_duplicates(subset=['pk'],keep='last') # 含更改後的訂單，不含更改前的訂單
    first = all_diff.drop_duplicates(subset=['pk'],keep='first') # 含更改前的訂單，不含更改後的訂單
    pk_diff = all_diff.drop_duplicates(subset=['pk'],keep=False) # 不含更改過的所有訂單

    # 找被更改過的資料
    new = last.append(pk_diff).drop_duplicates(subset=['pk'],keep=False) # 今天所有被更改過的資料
    old = first.append(pk_diff).drop_duplicates(subset=['pk'],keep=False) # 昨天所有被更改過的資料

    # 合併兩天所有被更改過的資料，並將兩天OEB16與TC_SFA104的數據分別列出
    merged_data = new.merge(old,on = ['pk','OEB01','OEB03','OEB15','OEA02','OEA14'], suffixes=('_new','_old') )
    # 取出交期被更改的資料
    changed_data = merged_data[merged_data['OEB16_new'] != merged_data['OEB16_old']].copy()
    
    # 預計完工日 = 預計下機日 + 10天
    date_of_completion = changed_data['TC_SFA104_new'] + datetime.timedelta(days = 10)
    # 找出可能來不及交貨的資料 (當預計完工日超過約定交貨日時)
    delayed_data = changed_data[date_of_completion > changed_data['OEB15']].copy()
        
    # 當沒有符合條件的訂單時，回傳None
    if delayed_data.empty == True:
        return None

    # 最後要寄給業務的通知內容
    final = delayed_data.loc[:, ['OEA14','OEA02','OEB01','OEB03','OEB15','OEB16_old','OEB16_new', 'TC_SFA104_new']]
    final.sort_values(['OEA14','OEB01'], inplace = True, ignore_index=True) # 依照業務編號、訂單編號做遞增排序
    final.rename(columns={'OEA02': '訂單日期', 'OEA14': '業務編號', 'OEB01': '訂單號碼', 'OEB03': '項次', 'OEB15': '約定交貨日', 'OEB16_old': '原排定交貨日', 'OEB16_new': '新排定交貨日', 'TC_SFA104_new': '預計下機日'}, inplace=True)
    final.index = final.index + 1 # index從1開始
    
    #dataframe轉html
    df_html = final.to_html(escape=False,index=True, justify = 'center')
    # 回傳final的html格式
    return df_html

def read_file(file):
    try:
        with open(file,'r',encoding="utf-8") as f:
            content_list = f.read().split('\n')
        return content_list
    except OSError as err:
        print("OS error: {0}".format(err))
    except:
        print("Unexpected error:", sys.exc_info()[0])
        raise

def SendMail(msg_html):
    #print('start to send mail...')
    emails = read_file('email.txt') #讀email資料
    sender = emails[0] #setup sender gmail,ex:"Fene1977@superrito.com"
    password = emails[1] #setup sender gmail password
    recipients = emails[2].split(', ') #setup recipients mail
    today_date = datetime.date.today() 
    sub = today_date.strftime("%m/%d") + "訂單交期更改通知" #step4:setup your subject
    
    outer = MIMEMultipart()
    outer['From'] = sender #setup sender gmail
    outer['To'] = ','.join(recipients) #setup recipient mail
    #outer["Cc"] = cc_mail #setup cc mail
    outer['Subject'] = sub #setup your subject

    #設定純文字資訊
    plainText = "偵測到訂單排定交貨日更改，且預計下機日與約定交貨日差距小於10天，請確認以下訂單："
    msgText = MIMEText(plainText, 'plain', 'utf-8')
    outer.attach(msgText)
    
    #設定HTML資訊
    htmlText = msg_html #step7:edit your mail content
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
        

def remove_file(file_name):
    try:
        os.remove(file_name)
    except OSError as e:
        print(e)
    else:
        print("Delete %s successfully."%(file_name))

# copy file1 as file2
def copy_file(file1, file2):
    try:
        shutil.copyfile(file1, file2)
        print("Copy %s as %s successfully."%(file1, file2))
    except OSError as err:
        print(err)

def main():
    #print('資料載入中...')
    fileNames = read_file('fileName.txt') #取得欲分析之檔案名稱
    yesterday = load_data(fileNames[0])
    today = load_data(fileNames[1])
    
    #print('資料處理中...')
    # 取得資料處理回傳結果
    for i in tqdm(range(0,1), ncols = 100, desc ="Progressing data..."):
        mail_content = data_process(yesterday, today)
    
    # 當回傳結果為None時，表示沒有符合條件的訂單；有回傳內容則以email寄出回傳內容
    if mail_content == None:
        print("沒有符合條件的訂單")
    else:
        print("找到需要通知的訂單")
        SendMail(mail_content)        
    
    copy_file(fileNames[1], fileNames[0]) # 將今天的檔案轉換為下次分析時昨天的檔案
    remove_file(fileNames[1]) #刪除今天的檔案

    input("請按任意鍵結束...") # 讓視窗停留，不要馬上關閉

main()
