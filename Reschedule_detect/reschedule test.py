#!/usr/bin/env python
# coding: utf-8

# In[1]:


import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText #專門傳送正文
import smtplib
import sys
import datetime
import os
from pandas.core.frame import DataFrame
import shutil


# In[2]:


#讀excel
def load_data(file_name, sheet = 0):
    df_data = pd.read_excel(file_name, sheet_name = sheet)
    return df_data


# In[3]:


'''資料處理跟抓取要放入郵件的內容
id_count: 計算目前在第幾個column
id[0]: column的index 
id[1]: 該column的value
id[1][n]: 該column的第n+1個row，因為n從0開始算
mail_list: 紀錄被改變的欄位值
'''
def data_process(df, df1):
    length = df.shape[0]
    id2_count = 0
    mail_list = []
    for id2 in df.iteritems(): 
        id2_count += 1 
        id1_count = 0
        for id1 in df1.iteritems():
            id1_count += 1
            # 對每個機台的訂單編號row做比對，直到最後一個row
            for i in range(1,length,7):
                #當兩個欄位不一樣且不在同一column時，記錄此訂單編號後六碼至mail_list。
                if(id2[1][i] == id1[1][i]) and (id2_count != id1_count): 
                    print("id1:",id1[1][i],"id2:",id2[1][i],"id1_count:",id1_count,"id2_count:",id2_count)
                    mail_list.append(id1[1][i])
                    
    # mail_list長度為0時，表示空list，即未修改
    if len(mail_list) == 0:
        return None
                   
    c = {"被變更的位置" : mail_list} #list轉df
    mail_list_df = DataFrame(c)
    
    #dataframe轉html
    df_html = mail_list_df.to_html(escape=False,index=False, justify = 'center')
    
    # 回傳final的html格式
    return df_html


# In[4]:


def read_file(file):
    with open(file ,'r',encoding="utf-8") as f:
            content_list = f.read().split('\n')
    return content_list


# In[5]:


def SendMail(msg_html, sender, password, recipients):

    today_date = datetime.date.today() 
    sub = today_date.strftime("%m/%d") + "訂單交期更改通知" #step4:setup your subject
    
    outer = MIMEMultipart()
    outer['From'] = sender #step:setup sender gmail
    outer['To'] = recipients #step:setup recipient mail
    #outer["Cc"] = cc_mail #step:setup cc mail
    outer['Subject'] = sub #step:setup your subject

    #設定純文字資訊
    plainText = "偵測到訂單排定交貨日更改，請確認以下訂單："
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




# In[7]:


import shutil
def file_copy_change_name(src, dst):
    try:
        shutil.copyfile(src, dst)
        print("Copy %s as %s successfully."%(src, dst))
    except OSError as err:
        print(err)



# In[8]:


def main():
    setup = read_file('UserGuide.txt') #取得欲分析之檔案名稱與Email
    yesterday = load_data(setup[0], setup[1])
    today = load_data(setup[2], setup[3])
    mail_content = data_process(yesterday, today)
    
    if mail_content == None:
        print("交期未更改")
    else:
        SendMail(mail_content, setup[4], setup[5], setup[6])
        print("交期已被更改")
    file_copy_change_name(setup[2], setup[0]) # 將今天的檔案轉換為下次分析時昨天的檔案
    input("請按任意鍵結束...") # 讓視窗停留，不要馬上關閉

main()

