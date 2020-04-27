from fuctions import *

path_xl = Path('tab.xlsx')  # Path(input('Введите путь до таблицы с данными: '))
contexts = readInfoFromBook(path_xl)

path_wd = Path('template.docx')  # Path(input('Введите путь до шаблона: '))
renderTpl(path_wd, contexts)
