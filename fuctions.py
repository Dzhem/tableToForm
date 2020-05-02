from pathlib import *
import xlrd
from docxtpl import DocxTemplate
from docx2pdf import convert


def readInfoFromBook(path):
    """
    :param path: Путь до xlsx файла с данными
    :return: список словарей, где ключ - название колонки, значение - значение ячейки по строкам
    """
    book = xlrd.open_workbook(path)
    sheet = book.sheet_by_index(0)
    keys = sheet.row_values(0)

    info = []
    for row_num in range(1, sheet.nrows):
        row = sheet.row_values(row_num)
        row_d = {}
        for ind, val in enumerate(row):
            row_d[str(keys[ind])] = val
        info.append(row_d)

    return info


def renderTpl(path_tpl, path_out, contexts, format='pdf'):
    """
    Заменяет значение полей шаблона данными из contexts и сохраняет новый файл
    :param format: формат конечных файлов (PDF / docx)
    :param path_tpl: Путь до docx-шаблона
    :param contexts: список словарей для замены полей шаблона
    """
    if path_out == '.':
        path_out = path_tpl.parent

    for i, context in enumerate(contexts):
        doc = DocxTemplate(path_tpl)
        doc.render(context)
        new_path = Path(f'{path_out}/{path_tpl.stem}_{i}{path_tpl.suffix}')
        doc.save(new_path)

        if format in ('pdf', 'PDF'):
            docxToPdf(new_path, path_out)


def docxToPdf(in_file, out_file):
    """
    Функция конвертирует docx-файл в pdf-файл с последующим удалением docx-файла
    :param in_file: путь к docx-файлу
    :param out_file: путь для сохранения pdf-файла
    :return:
    """
    convert(in_file, out_file)

    try:
        in_file.unlink()
    except OSError as e:
        print("Ошибка: %s : %s" % (in_file, e.strerror))
