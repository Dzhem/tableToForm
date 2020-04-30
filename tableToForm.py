from fuctions import *

path_xl = Path('tab.xlsx')  # Path(input('Введите путь до таблицы с данными: '))
path_wd = Path('template.docx')  # Path(input('Введите путь до шаблона: '))
format = input('Введите формат конечных файлов (PDF / docx): ')
format = format if format else 'pdf'

contexts = readInfoFromBook(path_xl)
renderTpl(path_wd, contexts, format)
