from fuctions import *
import PySimpleGUI as sg

sg.theme('LightGrey1')
layout = [[sg.Text('Таблица', size=(14, 1)), sg.InputText(), sg.FileBrowse('Обзор', file_types=(('Excel', '*.xlsx'),), size=(10, 1))],
          [sg.Text('Шаблон', size=(14, 1)), sg.InputText(), sg.FileBrowse(
              'Обзор', file_types=(('Word', '*.docx'),), size=(10, 1))],
          [sg.Text('Папка сохранения', size=(14, 1)), sg.InputText(),
           sg.FolderBrowse('Обзор', size=(10, 1))],
          [sg.Text('Формат сохранения', size=(14, 1)), sg.Radio('pdf', 'FORMAT', default=True), sg.Radio('docx', 'FORMAT'), sg.Submit('Старт!', size=(10, 1))]]

window = sg.Window('Table To Form', layout)

while True:
    event, values = window.read()

    if event in (None, 'Cancel'):
        break

    print('You entered ', values[0], values[1],
          values[2], values[3], values[4])
    path_xl = Path(values[0])
    path_wd = Path(values[1])

    if values[3]:
        format = 'pdf'
    else:
        format = 'docx'

    contexts = readInfoFromBook(path_xl)
    renderTpl(path_wd, contexts, format)

window.close()
