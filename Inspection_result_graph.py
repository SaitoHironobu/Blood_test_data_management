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




layout = [
    [sg.Text('各種血液グラフ表示アプリ')],
    [sg.Button('中性脂肪', key='-btn1-')]
]

window = sg.Window('各種血液グラフ表示アプリ', layout)

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

while True:
    event, value = window.read() 
    if event == '-btn1-':
        neutral_data = neutral_fat_value()
    elif event is None:
        break
    
window.close()



