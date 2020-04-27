from docxtpl import DocxTemplate
from fuctions import *


path_xl = Path(input('Введите путь до таблицы с данными: '))
contexts = readInfoFromBook(path_xl)


# doc = DocxTemplate('template.docx')

# TODO: реализовать чтение заголовков колонок,
#  элементы словаря context должны использовать эти заголовки
# for row_num in range(sheet.nrows):
#     row = sheet.row_values(row_num)
#
#     context = {}
#     for ind, val in enumerate(row):
#         context[str(f'col_{ind}')] = val
#     print(context)
#
#     doc.render(context)
#     doc.save(f'template{row_num}.docx')
