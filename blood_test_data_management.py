import PySimpleGUI as sg
import pandas as pd
from pathlib import Path
import random
import openpyxl


sg.SetOptions(text_justification='right')
layout = [
    [sg.Text('顧客情報登録フォーム',font=('小塚ゴシック pro B',16), pad=(50,0))],
    [sg.Text('会社名',size=(15,1)), sg.InputText(key='-Company-',size=(22,1))],
    [sg.Text('名前',size=(15,1)), sg.InputText(key='-Sei-', default_text= '性', size=(10,1)), sg.InputText(key='-Mei-', default_text= '名',size=(10,1))],
    [sg.Text('ふりがな',size=(15,1)), sg.InputText(key='-Kana_sei-', default_text= 'せい',size=(10,1)), sg.InputText(key='-Kana_mei-', default_text= 'めい',size=(10,1))],
    [sg.Text('メールアドレス',size=(15,1)), sg.InputText(key='-Email-', default_text= '@', size=(22,1))],
    [sg.Text('電話番号',size=(15,1)), sg.InputText(key='-Phone-', size=(22,1))],
    [sg.Text('性別',size=(15,1)), sg.Radio('男', 'sex', default= True, key='-Male-'),
    sg.Radio('女', 'sex',  key='-Female-')],
    [sg.Text('年齢',size=(15,1)), sg.Spin([i for i in range(20,90)], initial_value=35, key='-Age-')],
    [sg.Button('新規登録', key='-Register-', pad=(120,0), disabled=False)]
]

window = sg.Window('血液データ登録フォーム', layout)

def get_wgt_values():
    company = value['-Company-']
    sei = value['-Sei-']
    mei = value['-Mei-']
    kana_sei = value['-Kana_sei-']
    kana_mei = value['-Kana_mei-']
    email = value['-Email-']
    phone = value['-Phone-']
    if value['-Male-'] is True:
        sex = '男'
    else:
        sex = '女'
    age = value['-Age-']
    x = random.randrange(1000,9999)
    cust_id = email[:2] + str(x)
    data = [sei, mei, kana_sei, kana_mei, email, sex, age, phone, company, cust_id ]
    
    return data



def insert_to_db(data):
    try:
        wb = openpyxl.load_workbook('顧客情報管理台帳.xlsx')
        ws = wb['Sheet1']
        last_row = ws.max_row +1
        for colm in range(len(data)):
            ws.cell(row=last_row, column=colm + 1, value=data[colm])
        wb.save('顧客情報管理台帳.xlsx')
        status = 'succeed'
    except:
        status = None
    return status

def wgt_clear():
    window.find_element('-Company-').Update('')
    window.find_element('-Sei-').Update('')
    window.find_element('-Mei-').Update('')
    window.find_element('-Kana_sei-').Update('')
    window.find_element('-Kana_mei-').Update('')
    window.find_element('-Email-').Update('')
    window.find_element('-Phone-').Update('')
    window.find_element('-Age-').Update('')

while True:
    event, value = window.read() 
    if event == '-Register-':
        new_cust_data = get_wgt_values()

        result = insert_to_db(new_cust_data)
        
        if result == 'succeed':
            sg.popup('追記が成功しました。 ')
            wgt_clear()
        else:
            sg.popup('追記に失敗しました。 ')


    elif event is None:
        break
    
window.close()