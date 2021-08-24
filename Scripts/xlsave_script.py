import xlwt
from os import path


def xl_write(list, buys_data, y_coord):
    for buy in buys_data:
        list.write(y_coord, 0, buy['id'])
        list.write(y_coord, 1, buy['name'])
        list.write(y_coord, 2, buy['price'])
        list.write(y_coord, 3, buy['order'])
        list.write(y_coord, 4, buy['num'])
        list.write(y_coord, 5, buy['adr'])
        y_coord += 1
    return y_coord


def xl_create_book():
    book = xlwt.Workbook()
    sheet = book.add_sheet('MTSparse')
    sheet.write(0, 0, 'id Покупки')
    sheet.write(0, 1, 'Модель')
    sheet.write(0, 2, 'Цена')
    sheet.write(0, 3, 'Ордер')
    sheet.write(0, 4, 'Код покупки')
    sheet.write(0, 5, 'Адрес')
    y = 1
    print('book created')
    return book, sheet, y


def create_and_save_table(buys_data):
    file_path = path.abspath(path.join(path.dirname(__file__), '..', 'ЗаказыМТС.xls'))
    buys_book, buys_sheet, buys_y = xl_create_book()
    xl_write(buys_sheet, buys_data,  buys_y)
    buys_book.save(file_path)
