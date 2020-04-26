import xlrd
from docxtpl import DocxTemplate

# TODO: реализовать работу с путём к файлу
book = xlrd.open_workbook('tab.xlsx')
sheet = book.sheet_by_index(0)

doc = DocxTemplate('template.docx')

# TODO: реализовать чтение заголовков колонок,
#  элементы словаря context должны использовать эти заголовки
for row_num in range(sheet.nrows):
    row = sheet.row_values(row_num)

    context = {}
    for ind, val in enumerate(row):
        context[str(f'col_{ind}')] = val
    print(context)

    doc.render(context)
    doc.save(f'template{row_num}.docx')
