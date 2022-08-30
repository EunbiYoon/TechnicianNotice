import pandas as pd
import pandas as pd
import numpy as np
import xlrd
import smtplib
from email.message import EmailMessage
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from datetime import date
import msoffcrypto
import io
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import numpy as np
from email.mime.image import MIMEImage
import os
import matplotlib.pyplot as plt
from matplotlib import rc
from matplotlib.pyplot import figure
from datetime import date,timedelta
import datetime 
from dateutil.relativedelta import relativedelta
import smtplib # 메일을 보내기 위한 라이브러리 모듈
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication 




##########################데이터 정리하기
# 데이터 불러오기
data=pd.DataFrame()
file=['1_AGT Model.xlsx','2_Pro2 Model.xlsx','3_Plus Model.xlsx','4_Win Model.xlsx','5_CK 5.0 Model.xlsx']


# 오늘 날짜 추출
for i in range (2):
    data=pd.read_excel('//US-SO11-NA08765/R&D Secrets/M+3 task/200 Top Loader/1_Daily SVC Report/'+file[i],sheet_name='SVC')
    print(data)

    today=date.today()
    #today=date.today()-relativedelta(days=1)
    today=today.strftime('%Y-%m-%d') ##today의 형식을 바꾸겠다, string 변환

    #데이터 추출
    today_data=data[data['Report_Date'] == today]
    print(today_data.Technician)
    ex_today_data=today_data[today_data['Symptoms'] == 'EXPLANATION']
    is_sort=ex_today_data[['Technician','RCPT_NO_ORD_NO','DATA_TYPE']]
    is_sort=is_sort.dropna(axis=0)

    is_sort=is_sort.reset_index()
    is_sort=is_sort.drop(['index'],axis=1)
    tech_sort=is_sort['Technician'].str.lower()
    rnn_sort=is_sort['RCPT_NO_ORD_NO']
    email_sort=is_sort['DATA_TYPE']
    #print(rnn_sort)

    for i in range(len(is_sort)):
        k = tech_sort.loc[i].find(' ')
        #email_sort.loc[i]='eunbi1.yoon@lge.com'
        email_sort.loc[i]=tech_sort.loc[i][:k]+"."+tech_sort.loc[i][k+1:]+'@lge.com'

    final_sort=pd.concat([tech_sort,rnn_sort,email_sort],axis=1)
    final_sort.columns=["Tech","RNN","Email"]

    receiver=final_sort.Email.to_list()
    receiver=str(receiver).replace(r"', '",", ")
    receiver=str(receiver).replace(r"[","")
    receiver=str(receiver).replace(r"]","")
    receiver=str(receiver).replace(r"'","")
    print("Receiver")
    print(len(receiver))




    



