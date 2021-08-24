
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from time import sleep
from selenium.webdriver.common.keys import Keys
import re
import win32api
import win32con
from os import path
import cv2
import numpy as np
import pyscreenshot as ImageGrab
import pyautogui
import wx, ctypes

'''
addr_list = [
    "г. Москва, ул. Борисовские Пруды, 26",
    "г. Москва, ул. Митинская, 40",
    "г. Москва, ул. Константина Симонова, 2а",
    "п. Сосенское, ш. Калужское, 21 км, ТРЦ Мега Теплый Стан"
             ]

phones_order_dat = [
                    ["https://shop.mts.ru/product/smartfon-samsung-g970-galaxy-s10e-6-128gb-oniks", 2, 1],
                    ["https://shop.mts.ru/product/smartfon-xiaomi-mi-a3-4-64gb-black", 7, 2]
                    ]


name = "Василий"
second_name = "Пронин"
email = "fakemail@ya.ru"
phone_number = "9857411742"
'''

'''
def find_patt(image, patt, thres):
  img_grey = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
  patt_H, patt_W = patt.shape[:2]
  res = cv2.matchTemplate(img_grey, patt, cv2.TM_CCOEFF_NORMED)
  loc = np.where(res > thres)
  return patt_H, patt_W, list(zip(*loc[::-1]))
'''


def click(x, y):
    win32api.SetCursorPos((x, y))
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, x, y, 0, 0)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, x, y, 0, 0)


def click_submit_store_btn(address_form, driver):
    '''
    app = wx.PySimpleApp()

    SM_XVIRTUALSCREEN = 76
    SM_YVIRTUALSCREEN = 77
    SM_CXVIRTUALSCREEN = 78
    SM_CYVIRTUALSCREEN = 79

    user32 = ctypes.windll.user32
    width, height = user32.GetSystemMetrics(SM_CXVIRTUALSCREEN), user32.GetSystemMetrics(SM_CYVIRTUALSCREEN)
    x, y = user32.GetSystemMetrics(SM_XVIRTUALSCREEN), user32.GetSystemMetrics(SM_YVIRTUALSCREEN)

    screen = wx.ScreenDC()
    bmp = wx.EmptyBitmap(width, height)
    mem = wx.MemoryDC(bmp)
    mem.Blit(0, 0, width, height, screen, x, y)
    del mem
    bmp.SaveFile('screenshot.png', wx.BITMAP_TYPE_PNG)

    screenshot = ImageGrab.grab()
    img = np.array(screenshot.getdata(), dtype='uint8').reshape((screenshot.size[1], screenshot.size[0], 3))

    patt = cv2.imread('../button.bmp', 0)
    h, w, points = find_patt(img, patt, 0.60)
    if len(points) != 0:
        pyautogui.moveTo(points[0][0] + w / 2, points[0][1] + h / 2)
        pyautogui.click()
    '''

    button = address_form.find_element(By.CLASS_NAME, "btn")
    window_pos = driver.get_window_position()
    wn_pos_x = window_pos["x"]
    wn_pos_y = window_pos["y"]
    location = button.location
    x = location["x"]
    y = location["y"]
    while True:
        click(wn_pos_x + x + 150, wn_pos_y + y + 200)
        try:
            driver.find_element(By.CLASS_NAME, "selectedStore")
            break
        except:
            y = y-1
            pass


def input_person_info(driver, By_obj, class_name, value):
    form = driver.find_element(By_obj, class_name)
    for _ in range(25):
        form.send_keys(Keys.BACK_SPACE)
    form.send_keys(value)


def add_item_to_basket(driver, link):
    driver.get(link)
    sleep(2)
    try:
        driver.find_element(By.CLASS_NAME, "PushTip-close").click()
    except:
        pass
    is_button_found = False
    i = 0
    while not is_button_found:
        try:
            driver.find_element(By.CLASS_NAME, "buybutton").click()
            is_button_found = True
        except:
            driver.get(link)
            i += 1
            if i > 15:
                is_button_found = True
    sleep(1)


def optioning_basket(driver):
    driver.get("https://shop.mts.ru/personal/basket")
    telephone_name = driver.find_element(By.XPATH, "//*[@id=\"basket_form cart-form\"]"
                                                   "/div[3]/div[1]/div/div[2]/div[1]/div[2]/div/a/h3").text
    sleep(1)
    try:
        additional_item = driver.find_element(By.CLASS_NAME, "additional")
        is_additional_item = True
    except:
        is_additional_item = False
    while is_additional_item:
        try:
            additional_item = driver.find_element(By.CLASS_NAME, "additional")
            sleep(1)
            additional_item.find_element(By.CLASS_NAME, "hint-close").click()
            sleep(3)
            driver.get("https://shop.mts.ru/personal/basket")
        except:
            is_additional_item = False
    price = driver.find_element(By.CLASS_NAME, "total")
    price = price.find_element(By.CLASS_NAME, "number").text
    sleep(1)
    driver.find_element(By.ID, "submitBasket").click()
    return telephone_name, price


def input_delivery_point(driver, address_list):
    sleep(2)
    ad_list = address_list.copy()
    for address in address_list:
        address_form = driver.find_element(By.CLASS_NAME, "pickup__container")
        try:
            driver.find_element(By.CLASS_NAME, "point__back-link")
            sleep(1)
        except:
            pass
        try:
            address_input = driver.find_element(By.XPATH,
                                                "//*[@id=\"IM_FORM\"]/div[1]/div/fieldset[2]/div[2]/div[2]/div[1]/div[2]/div[1]/div[2]/input")
        except:
            address_input = driver.find_element(By.XPATH,
                                                "//*[@id=\"IM_FORM\"]/div[1]/div/fieldset[3]/div[2]/div[2]/div/div[2]/div[1]/div[2]/input")
        is_sallable_this_address = True
        for _ in range(50):
            address_input.send_keys(Keys.BACK_SPACE)
        sleep(1)
        address_input.send_keys(address)
        sleep(5)
        try:
            driver.find_element(By.CLASS_NAME, "warning__not-found")
            is_sallable_this_address = False
        except:
            pass
        if is_sallable_this_address:
            driver.find_element(By.CLASS_NAME, "store__row").click()
            sleep(1)
            click_submit_store_btn(address_form, driver)
            sleep(1)
            try:
                driver.find_element(By.ID, "make_order").click()
                sleep(6)
                order_text = driver.find_element(By.CLASS_NAME, "contents").text
                order_num = re.search('[0-9]+', order_text)[0]
                return {order_num: ['------', address]}, ad_list
            except:
                pass
        ad_list.remove(address)
    return None, ad_list


def make_order(driver, address_list, name, last_name, email, personal_phone):
    input_person_info(driver, By.ID, "order_contact_NAME", name)
    input_person_info(driver, By.ID, "order_contact_LAST_NAME", last_name)
    input_person_info(driver, By.ID, "order_contact_PERSONAL_PHONE", personal_phone)
    input_person_info(driver, By.ID, "order_contact_EMAIL", email)
    order_dat, new_addr_list = input_delivery_point(driver, address_list)  # address, order_num
    sleep(5)
    return order_dat, new_addr_list


def delete_item_from_busket(driver):
    driver.get("https://shop.mts.ru/personal/basket")
    is_basket_empty = False
    sleep(1)
    while not is_basket_empty:
        try:
            driver.find_element(By.XPATH, "//*[@id=\"basket_form cart-form\"]/div[3]/div[1]/div/div[2]/a").click()
        except:
            is_basket_empty = True


def initializate_web_driver():
    options = Options()
    options.add_argument('--disable-features=NetworkService')
    options.add_argument("--disable-notifications")
    options.add_argument('--enable-features=NetworkServiceWindowsSandbox')
    options.add_argument("--disable-infobars")
    chrome_path = path.abspath(path.join(path.dirname(__file__), 'chromedriver.exe'))
    driver = webdriver.Chrome(chrome_path, options=options)
    return driver


def parse_script_main(addr_list, phones_order_dat, user_dat):
    result = {}
    if user_dat[0]:
        name = user_dat[0]
    else:
        name = 'Имя'
    if user_dat[1]:
        second_name = user_dat[1]
    else:
        second_name = 'Фамилия'
    if user_dat[2]:
        email = user_dat[2]
    else:
        email = 'fakemail@ya.ru'
    if user_dat[3]:
        phone_number = user_dat[3]
    else:
        phone_number = '9857411743'

    driver = initializate_web_driver()

    try:
        for phone in phones_order_dat:
            try:
                res_order_dat = {}
                phone_id = phone[2]
                phone_link = phone[0]
                count = int(phone[1])
                phone_name = ''
                price = ''
                is_sellable = True

                addr_list_copy = addr_list.copy()
                i = 0
                try:
                    while i < count and is_sellable:
                        add_item_to_basket(driver, phone_link)
                        phone_name, price = optioning_basket(driver)
                        # order_dat = [address, order_num]
                        order_dat, addr_list_copy = make_order(driver, addr_list_copy, name, second_name, email, phone_number)
                        delete_item_from_busket(driver)

                        if not order_dat:
                            is_sellable = False
                        else:
                            res_order_dat.update(order_dat)
                        i += 1
                except:
                    phone_id = phone[2]
                    result.update({phone_id: ['Ошибка', '-', {'-': ['-', '-']}]})
                result.update({phone_id: [phone_name, price, res_order_dat]})
                delete_item_from_busket(driver)
            except:
                phone_id = phone[2]
                result.update({phone_id: ['Ошибка', '-', {'-': ['-', '-']}]})
    except:
        for phone in phones_order_dat:
            phone_id = phone[2]
            # {Айди телефона: [Модель, Цена, {Ордер: [Код продажи, Адрес] } ] }
            result.update({phone_id: ['Ошибка', '-', {'-': ['-', '-']}]})
        driver.close()
        return result
    driver.close()
    return result
