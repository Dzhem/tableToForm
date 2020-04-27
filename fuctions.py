from pathlib import *
import xlrd


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