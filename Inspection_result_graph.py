#import seaborn
import pandas as pd
import numpy as np
import matplotlib.style
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import openpyxl
from pathlib import Path
import matplotlib.ticker as ticker
import PySimpleGUI as sg



#sg.theme('DarkAmber') 
sg.theme('Topanga') 
layout = [
    [sg.Text('各種血液グラフ表示アプリ')],
    [sg.Button('中性脂肪', key='-btn2-')],
    [sg.Button('LDLコレステロール', key='-btn3-')],
    [sg.Button('尿素窒素(BUN)', key='-btn4-')],
    [sg.Button('クレアチニン', key='-btn5-')],
    [sg.Button('尿酸(UA)', key='-btn6-')],
    [sg.Button('血糖(グルコース)', key='-btn7-')],
    [sg.Button('HbA1c(NGSP)', key='-btn8-')],
    [sg.Button('体重', key='-btn9-')]
]

window = sg.Window('bgdisplay', layout)

#中性脂肪ボタン
def neutral_fat_value():
    path = Path('血液検査データ.xlsx')
    df = pd.read_excel(path, sheet_name='Sheet1', header=0, engine='openpyxl')
    x0 = df.loc[:, '検査日'].tolist()
    x1 =  df['中性脂肪\n(TG)'].tolist()


    fig,ax = plt.subplots()
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%Y/%m'))

    x =(x0) 
    y = (x1) 

    plt.plot(x, y)
    plt.show()
    
#LDLコレステロール
def neutral_ldl_value():
    path = Path('血液検査データ.xlsx')
    df = pd.read_excel(path, sheet_name='Sheet1', header=0, engine='openpyxl')
    x0 = df.loc[:, '検査日'].tolist()
    x1 =  df['LDLコレステロール'].tolist()


    fig,ax = plt.subplots()
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%Y/%m'))

    x =(x0) 
    y = (x1) 

    plt.plot(x, y)
    plt.show()
    
    #尿素窒素
def urea_nitrogen_value():
    path = Path('血液検査データ.xlsx')
    df = pd.read_excel(path, sheet_name='Sheet1', header=0, engine='openpyxl')
    x0 = df.loc[:, '検査日'].tolist()
    x1 =  df['尿素窒素\n(BUN)'].tolist()


    fig,ax = plt.subplots()
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%Y/%m'))

    x =(x0) 
    y = (x1) 

    plt.plot(x, y)
    plt.show()
    
    #クレアチニン
def urea_nitrogen_value():
    path = Path('血液検査データ.xlsx')
    df = pd.read_excel(path, sheet_name='Sheet1', header=0, engine='openpyxl')
    x0 = df.loc[:, '検査日'].tolist()
    x1 =  df['クレアチニン'].tolist()


    fig,ax = plt.subplots()
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%Y/%m'))

    x =(x0) 
    y = (x1) 

    plt.plot(x, y)
    plt.show()

#尿酸 uric acid
def uric_acid_value():
    path = Path('血液検査データ.xlsx')
    df = pd.read_excel(path, sheet_name='Sheet1', header=0, engine='openpyxl')
    x0 = df.loc[:, '検査日'].tolist()
    x1 =  df['尿酸\n(UA)'].tolist()


    fig,ax = plt.subplots()
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%Y/%m'))

    x =(x0) 
    y = (x1) 

    plt.plot(x, y)
    plt.show()
    
    
#血糖(グルコース)Blood sugar
def blood_sugar_value():
    path = Path('血液検査データ.xlsx')
    df = pd.read_excel(path, sheet_name='Sheet1', header=0, engine='openpyxl')
    x0 = df.loc[:, '検査日'].tolist()
    x1 =  df['血糖\n(グルコース)'].tolist()


    fig,ax = plt.subplots()
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%Y/%m'))

    x =(x0) 
    y = (x1) 

    plt.plot(x, y)
    plt.show()


#HbA1c(NGSP)
def hba1c_value():
    path = Path('血液検査データ.xlsx')
    df = pd.read_excel(path, sheet_name='Sheet1', header=0, engine='openpyxl')
    x0 = df.loc[:, '検査日'].tolist()
    x1 =  df['HbA1c\n(NGSP)'].tolist()


    fig,ax = plt.subplots()
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%Y/%m'))

    x =(x0) 
    y = (x1) 

    plt.plot(x, y)
    plt.show()

#体重 body weight 
def body_weight_value():
    path = Path('血液検査データ.xlsx')
    df = pd.read_excel(path, sheet_name='Sheet1', header=0, engine='openpyxl')
    x0 = df.loc[:, '検査日'].tolist()
    x1 =  df['体重'].tolist()


    fig,ax = plt.subplots()
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%Y/%m'))

    x =(x0) 
    y = (x1) 

    plt.plot(x, y)
    plt.show()




while True:
    event, value = window.read() 
    if event == '-btn2-':
        neutral_data = neutral_fat_value()
        
    elif event == '-btn3-':
        neutral_data02 = neutral_ldl_value()
        
    elif event == '-btn4-':
        neutral_data03 = urea_nitrogen_value()
        
    elif event == '-btn5-':
        neutral_data04 = urea_nitrogen_value()
        
    elif event == '-btn6-':
        netural_data05 = uric_acid_value()
        
    elif event == '-btn7-':
        netural_data07 = blood_sugar_value()
        
    elif event == '-btn8-':
        netural_data08 = hba1c_value()
        
    elif event == '-btn9-':
        body_weight =  body_weight_value()
        
    elif event is None:
        break
    
window.close()



