import PySimpleGUI as sg
import pandas as pd

import openpyxl

sg.theme('DarkAmber')

EXCEL_FILE = 'reg.xlsx'
df = pd.read_excel(EXCEL_FILE)
layout = [
    [sg.Text('Please fill out the following fields:')],
    [sg.Text('Name', size=(15, 1)), sg.InputText(key='Name')],
    [sg.Text('Phone', size=(15, 1)), sg.InputText(key='Phone')],
    [sg.Text('Gender', size=(15, 1)), sg.Combo(['Male', 'Female', 'Other'], key='Gender')],
    [sg.Text('Email', size=(15, 1)), sg.InputText(key='Email')],
    [sg.Text('Password', size=(15, 1)), sg.InputText(key='Password')],
    [sg.Text(' Re-Enter Password', size=(15, 1)), sg.InputText(key='Re-Enter Password')],
    [sg.Submit(), sg.Button('Clear'), sg.Exit()]
]

window = sg.Window('Registration', layout)


def clear_input():
    for key in values:
        window[key]('')
        return None


while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Exit':
        break
    if event == 'Clear':
        clear_input()
    if event == 'Submit':
        df = df.append(values, ignore_index=True)
        df.to_excel(EXCEL_FILE, index=False)
        sg.popup('Registered')
        clear_input()
        window.close()
