import sys
import re
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *
import imap_sript
import xlsave_script
import parse_script
from os import path

# import win32gui, win32con

# The_program_to_hide = win32gui.GetForegroundWindow()
# win32gui.ShowWindow(The_program_to_hide , win32con.SW_HIDE)


def take_data(file_path):
    file_path = path.abspath(path.join(path.dirname(__file__), '..', 'Options', file_path))
    result = []
    file = open(file_path, 'r', encoding="utf-8")
    for line in file:
        result.append(line)
    return result


reg = {}
google_imap_srv = 'imap.gmail.com'
addr_list = take_data(r'addresses.txt')
user_dat = take_data(r'user.txt')


class MainWindow(QWidget):

    def __init__(self):
        super().__init__()
        self.main_layout = QVBoxLayout()
        self.contentLayout = QVBoxLayout()
        self.initUI()

    def initUI(self):
        self.setWindowTitle('МТС парсер')

        btns_layout = QHBoxLayout()

        save_dat_btn = QPushButton('Сохранить')
        save_dat_btn.setFixedWidth(120)
        save_dat_btn.clicked.connect(self.saveData)

        create_exel_btn = QPushButton('Создать таблицу')
        create_exel_btn.setFixedWidth(120)
        create_exel_btn.clicked.connect(self.createExelFile)

        add_phone_btn = QPushButton('Добавить телефон')
        add_phone_btn.setFixedWidth(120)
        add_phone_btn.clicked.connect(self.addPhoneLine)

        make_orders_btn = QPushButton('Заказать')
        make_orders_btn.setFixedWidth(120)
        make_orders_btn.clicked.connect(self.makeOrder)

        find_buy_orders_btn = QPushButton('Проверить коды покупки')
        find_buy_orders_btn.setFixedWidth(140)
        find_buy_orders_btn.clicked.connect(self.getSellOrders)

        btns_layout.addWidget(add_phone_btn)
        btns_layout.addWidget(make_orders_btn)
        btns_layout.addWidget(find_buy_orders_btn)
        btns_layout.addWidget(create_exel_btn)
        btns_layout.addWidget(save_dat_btn)

        self.main_layout.addLayout(btns_layout)
        self.main_layout.addLayout(self.contentLayout)

        self.setLayout(self.main_layout)

        self.loadData()

        self.show()

    def addPhoneLine(self):
        if reg:
            line_id = max(reg.keys()) + 1
        else:
            line_id = 1
        phone_line = phoneLine(line_id)
        reg.update({line_id: phone_line})
        self.contentLayout.addLayout(phone_line)

    def makeOrder(self):
        phone_dat = []
        for key in reg.keys():
            if not reg[key].purchased:
                phone_dat.append([reg[key].link_edit.text(), reg[key].count_edit.text(),  reg[key].phone_id])
                reg[key].link_edit.setReadOnly(True)
                reg[key].link_edit.setStyleSheet("""QLineEdit { background-color: #D3D3D3}""")
                reg[key].count_edit.setReadOnly(True)
                reg[key].count_edit.setStyleSheet("""QLineEdit { background-color: #D3D3D3}""")
                reg[key].purchased = True
        parsed_data = parse_script.parse_script_main(addr_list, phone_dat, user_dat)
        # parsed_data = {1: ['Смартфон Apple iPhone 6s 32GB Space Gray', '24 990', {'24037090': ['------', 'г. Москва, ул. Борисовские Пруды, 26']}], 2: ['Смартфон Samsung A705 Galaxy A70 6/128Gb Black', '29 990', {'24037138': ['------', 'п. Сосенское, ш. Калужское, 21 км, ТРЦ Мега Теплый Стан'], '24037216': ['------', 'п. Сосенское, ш. Калужское, 21 км, ТРЦ Мега Теплый Стан']}], 3: ['Смартфон Honor 8X 64 Gb Blue', '14 990', {'24037264': ['------', 'п. Сосенское, ш. Калужское, 21 км, ТРЦ Мега Теплый Стан'], '24037324': ['------', 'п. Сосенское, ш. Калужское, 21 км, ТРЦ Мега Теплый Стан'], '24037498': ['------', 'п. Сосенское, ш. Калужское, 21 км, ТРЦ Мега Теплый Стан']}]}
        if parsed_data:
            for key in parsed_data.keys():
                reg[key].phone_name = parsed_data[key][0]
                reg[key].name_lbl.setText(parsed_data[key][0])
                if parsed_data[key][0] == 'Ошибка':
                    reg[key].purchased = False
                else:
                    reg[key].purchased = True
                reg[key].phone_price = parsed_data[key][1]
                reg[key].price_lbl.setText(parsed_data[key][1])
                reg[key].applyOrders(parsed_data[key][2])


        # idТелефона: [ордер1, ордер2 ...]
    def getSellOrders(self):
        request = {}
        for key in reg.keys():
            if reg[key].purchased:
                orders = []
                for order in reg[key].order_nums.keys():
                    orders.append(order)
                request.update({key: orders})
            else:
                pass

        # imap_script_main(login, password, imap_srv, order_num_list, search_numer):
        # {idТелефона: {ордер1: байкод1, ордер2: байкод2}, ...}
        response = imap_sript.imap_script_main(user_dat[2], user_dat[4], google_imap_srv, request)
        # response = {2: {'24032212': '4123'}, 3: {'24032230': '1234', '24032266': '2212'}, 5: {'24032308': '6341', '24032374': '2334', '24032422': '2316'}}
        for key in response.keys():
            reg[key].applyBuyOrders(response[key])

    # id, name, linkm pricem orderm num, adr
    def createExelFile(self):
        xl_wt = []
        for key in reg.keys():
            for order in reg[key].order_nums.keys():
                xl_wt.append({'id': key,
                              'name': reg[key].phone_name,
                              'link': reg[key].link_edit.text(),
                              'price': reg[key].phone_price,
                              'order': order,
                              'num': reg[key].order_nums[order][0],
                              'adr': reg[key].order_nums[order][1]})
        xlsave_script.create_and_save_table(xl_wt)

    def saveData(self):
        save_file_path = path.abspath(path.join(path.dirname(__file__), '..', 'Save', 'save'))
        save_file = open(save_file_path, 'w')

        for key in reg.keys():
            file_write(save_file, key)
            file_write(save_file, reg[key].link_edit.text())
            file_write(save_file, reg[key].count_edit.text())
            file_write(save_file, reg[key].phone_name)
            file_write(save_file, reg[key].phone_price)
            file_write(save_file, reg[key].purchased)
            file_write(save_file, str(len(reg[key].order_nums)))
            for ord_key in reg[key].order_nums.keys():
                order = str(ord_key) + '=!=' + reg[key].order_nums[ord_key][0] + '=!=' + reg[key].order_nums[ord_key][1]
                file_write(save_file, order)

        save_file.close()

    def loadData(self):
        save_file_path = path.abspath(path.join(path.dirname(__file__), '..', 'Save', 'save'))
        save_file = open(save_file_path, 'r')

        save_file_array = []
        for line in save_file:
            save_file_array.append(re.split('\n', line)[0])

        i = 0
        while i < len(save_file_array):
            phone_id = int(save_file_array[i])
            i += 1
            link = save_file_array[i]
            i += 1
            count = save_file_array[i]
            i += 1
            name = save_file_array[i]
            i += 1
            price = save_file_array[i]
            i += 1
            if save_file_array[i] == 'True':
                purchased = True
            else:
                purchased = False
            i += 1
            j = int(save_file_array[i])
            order = {}
            while j > 0:
                i += 1
                ord_dat = re.split(r'=!=', save_file_array[i])
                order.update({ord_dat[0]: [ord_dat[1], ord_dat[2]]})
                j -= 1
            i += 1
            phone_line = phoneLine(phone_id, link=link, count=count, phone_name=name, phone_price=price, order_nums=order, purchased=purchased)
            reg.update({phone_id: phone_line})
            self.contentLayout.addLayout(phone_line)

        save_file.close()


def file_write(file, value):
    if value:
        value = str(value)
        value.strip('\n')
        file.write(str(value))
        file.write('\n')
    else:
        file.write('-')
        file.write('\n')


class phoneLine(QHBoxLayout):
    def __init__(self,
                 phone_id,
                 purchased=False,
                 order_nums={},
                 phone_name='Модель телефона',
                 phone_price='Цена',
                 link='',
                 count=''
                 ):
        super().__init__()

        # variables
        self.order_nums = order_nums
        self.phone_id = phone_id
        self.purchased = purchased
        self.phone_name = phone_name
        self.phone_price = phone_price

        # Qt_objects
        self.link_edit = QLineEdit()
        self.count_edit = QLineEdit()
        self.delete_btn = QPushButton()
        self.name_lbl = QLabel(phone_name)
        self.price_lbl = QLabel(phone_price)
        self.order_nums_layout = QVBoxLayout()

        self.link_edit.setText(link)
        self.count_edit.setText(count)

        self.initUI()

    def initUI(self):

        self.link_edit.setPlaceholderText('Ссылка на телефон')
        self.count_edit.setPlaceholderText('Кол-во')
        self.link_edit.setFixedWidth(500)
        self.count_edit.setFixedWidth(50)
        self.delete_btn.setFixedWidth(120)
        self.name_lbl.setFixedWidth(300)
        self.price_lbl.setFixedWidth(53)

        regular = QRegExp("[0-9]{4}")
        pIntValidator = QRegExpValidator(self)
        pIntValidator.setRegExp(regular)
        self.count_edit.setValidator(pIntValidator)

        self.delete_btn.setText('Удалить телефон')
        self.delete_btn.clicked.connect(self.deletePhoneLine)

        self.applyOrders(self.order_nums)

        self.addWidget(self.link_edit)
        self.addWidget(self.count_edit)
        self.addWidget(self.delete_btn)
        self.addWidget(self.name_lbl)
        self.addWidget(self.price_lbl)

        if self.purchased:
            self.link_edit.setReadOnly(True)
            self.link_edit.setStyleSheet("""QLineEdit { background-color: #D3D3D3}""")
            self.count_edit.setReadOnly(True)
            self.count_edit.setStyleSheet("""QLineEdit { background-color: #D3D3D3}""")
            self.purchased = True

        self.addLayout(self.order_nums_layout)

    def deletePhoneLine(self):
        for i in reversed(range(self.order_nums_layout.count())):
            self.order_nums_layout.itemAt(i).widget().setParent(None)
        self.order_nums_layout.setParent(None)
        for i in reversed(range(self.count())):
            self.itemAt(i).widget().setParent(None)
        del reg[self.phone_id]
        self.deleteLater()

    def rewriteLabels(self):
        for i in reversed(range(self.order_nums_layout.count())):
            self.order_nums_layout.itemAt(i).widget().setParent(None)
        for key in self.order_nums.keys():
            lbl = QLabel('№{}: код продажи №{} Адрес: {}'.format(key, self.order_nums[key][0], self.order_nums[key][1]))
            self.order_nums_layout.addWidget(lbl)

    def applyOrders(self, dictionary):
        self.order_nums = dictionary
        self.rewriteLabels()

    def applyBuyOrders(self, dictionary):
        for key in dictionary:
            self.order_nums[key][0] = dictionary[key]
        self.rewriteLabels()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = MainWindow()
    sys.exit(app.exec_())



