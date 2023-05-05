from openpyxl import load_workbook
import PySimpleGUI as sg
from datetime import datetime

layout = [[sg.Text('Название мероприятия'),
           sg.Push(), sg.Input(key='Event')],
          [sg.Text('Фамилия'), sg.Push(), sg.Input(key='name2')],
          [sg.Text('Имя'), sg.Push(), sg.Input(key='name')],
          [sg.Text('Отчество'), sg.Push(), sg.Input(key='name3')],
          [sg.Text('Mail-почта'), sg.Push(), sg.Input(key='mail')],
          [sg.Text('номер телефона'), sg.Push(), sg.Input(key='NumberOfPhone')],
          [sg.Text('Тип билета'), sg.Push(), sg.Input(key='TypeOfTicket')],
          [sg.Text('Сектор'), sg.Push(), sg.Input(key='Sector')],
          [sg.Text('Ряд'), sg.Push(), sg.Input(key='Row')],
          [sg.Text('место'), sg.Push(), sg.Input(key='Place')],
          [sg.Button('Добавить'), sg.Button('Закрыть')]]

window = sg.Window('Data Entry', layout, element_justification='center')

while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Закрыть':
        break
    if event == 'Добавить':
        try:
            wb = load_workbook('sport.xlsx')
            sheet = wb['Лист1']
            ID = len(sheet['ID'])
            time_stamp = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
            data = [ID, values['Event'], values['name2'], values['name'], values['name3'], values['mail'], values['NumberOfPhone'], values['TypeOfTicket'], values['Sector'], values['Row'], values['Place'], time_stamp]
            sheet.append(data)
            wb.save('sport.xlsx')

            window['Event'].update(value='')
            window['name2'].update(value='')
            window['name'].update(value='')
            window['name3'].update(value='')
            window['mail'].update(value='')
            window['NumberOfPhone'].update(value='')
            window['TypeOfTicket'].update(value='')
            window['Sector'].update(value='')
            window['Row'].update(value='')
            window['Place'].update(value='')
            window['Event'].set_focus()
            sg.popup('Данные сохранены')
        except PermissionError:
            sg.popup('File in use', 'File is being used by another User.\nPlease try again later.')
window.close()