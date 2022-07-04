from curses import window
import PySimpleGUI as sg
import pandas as pd

#add some color to the window
sg.theme('DarkTeal9')


excel_file = 'excel_file.xlsx'
df = pd.read_excel(excel_file)

layout = [
    [sg.Text('Please Fill Out the Fields:')],
     [sg.Text('Full Name: ', size=(15,1)), sg.InputText(key = 'Name')],
    [sg.Text('MatricNo:', size=(10,1)), sg.InputText(key='Matric-Number')],
    [sg.Text('Level:', size=(6,1)), sg.Combo(['100L', '200L', '300L', '400L', '500L', '600L'], key='Level')],
    [sg.Text('Department:', size=(25,1)), sg.InputText(key='Department')],
    [sg.Text('CGPA:', size=(25,1)), sg.InputText(key='CGPA')],
   
    [sg.Text('I speak', size=(15,1)),
                            sg.Checkbox('German', key='German'),
                            sg.Checkbox('Spanish', key='Spanish'),
                            sg.Checkbox('French', key='French')],
    [sg.Text('Number of courses offered this semester:'), sg.Spin([i for i in range (0,10)],
                                                                    initial_value=0, key='Courses-Offered')],
    [sg.Submit(), sg.Button('Clear'), sg.Exit()]
]

window = sg.Window('Lead City Students Records System', layout)

def clear_input():
    for key in values:
        window['key']('')
    return None


while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Exit':
        break

    if event  == 'Clear':
        clear_input()

    if event == 'Submit':
        df =  df.append(values, ignore_index=True)
        df.to_excel(excel_file, index=False)
        sg.popup('Data Saved!')
        clear_input()
window.close()