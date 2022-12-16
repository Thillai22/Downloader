# -*- coding: utf-8 -*-
"""
Created on Fri Dec 24 22:58:00 2021

@author: ThillaiNathan
"""

from PIL import Image
from pytesseract import image_to_string
import time
import os
import shutil
import pyautogui
from time import process_time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By


start = process_time()
s = Service('C:/Users/thill/Downloads/chromedriver.exe')
driver = webdriver.Chrome(service=s)

# Defining the global variables
block_number_list = []
t, k, w = 0, 0, 0
survey_number, sub_division_number, block_number = 1, 1, 1


def browser():
    driver.get('https://eservices.tn.gov.in/eservicesnew/land/chittaCheckNewUrban_en.html?lan=en')
    driver.maximize_window()


def captcha():
    driver.save_screenshot('screenshot.png')
    im = Image.open('screenshot.png')
    im = im.crop((330, 770, 600, 850))  # defines crop points
    im.save('screenshot.png')
    captcha_text = image_to_string(Image.open('screenshot.png'))
    box = driver.find_element(By.XPATH, '//*[@id="captcha"]')
    box.send_keys(captcha_text.strip())


def xpath():
    district = Select(driver.find_element(By.XPATH, '//*[@id="districtCode"]'))
    taluk = Select(driver.find_element(By.XPATH, '//*[@id="auto_off"]/div[8]/div[2]/select'))
    town = Select(driver.find_element(By.XPATH, '//*[@id="auto_off"]/div[11]/div[2]/select'))
    ward = Select(driver.find_element(By.XPATH, '//*[@id="auto_off"]/div[14]/div[2]/select'))

    district.select_by_value('02')  # Chennai
    driver.implicitly_wait(10)
    taluk.select_by_value('14')  # Aminjikarai
    town.select_by_value('001')  # Villivakkam
    ward.select_by_value('005')  # Ward-001


def submit():
    submit_btn = driver.find_element(By.XPATH, '//*[@id="auto_off"]/div[30]/div[2]/input')
    submit_btn.click()


def download(sno, subno):
    time.sleep(1)
    pyautogui.hotkey('ctrl', 'p')
    time.sleep(2)
    pyautogui.press('enter')
    time.sleep(2)
    pyautogui.typewrite('SNo_' + check(sno) + '  SubDivNo_' + check(subno))
    time.sleep(2)
    pyautogui.press('enter')


def print_func():
    pyautogui.hotkey('ctrl', 'p')
    time.sleep(2)
    pyautogui.click(1417, 192)
    time.sleep(1)
    pyautogui.click(1438, 240)
    time.sleep(1)
    pyautogui.click(1438, 328)
    pyautogui.click(1377, 377)
    pyautogui.click(1547, 898)


def check(val_ele):
    val_ele = val_ele.replace('/', '-')
    return val_ele


def repeat(value, elem, sub):
    browser()
    xpath()
    captcha()
    time.sleep(3)
    block_func()
    block_number.select_by_value(value)
    survey_func()
    survey_number.select_by_value(elem)
    sub_division_func()
    sub_division_number.select_by_value(sub)
    submit()


def block_func():
    global block_number
    block_number = Select(driver.find_element(By.XPATH, '//*[@id="auto_off"]/div[15]/div[2]/select'))


def survey_func():
    global survey_number
    survey_number = Select(driver.find_element(By.NAME, 'surveyNo'))


def sub_division_func():
    global sub_division_number
    sub_division_number = Select(driver.find_element(By.NAME, 'subdivNo'))


def move(path):
    ent_dir_path = r"C:\Users\thill\Downloads"  # source directory        Change these two directories
    out_dir_path = r"D:\Pdf\District-Chennai\Taluk-Ambattur\Town-Ambattur\Ward-9 (059) Mogappair (WardH)\Block-" + path  # destination directory

    if not os.path.exists(out_dir_path):  # create folder if it doesn't exist
        os.makedirs(out_dir_path)

    file_type = "SNo_"

    entries = os.listdir(ent_dir_path)

    for entry in entries:
        if entry.startswith(file_type):
            # print(ent_dir_path + entry)
            shutil.move(os.path.join(ent_dir_path, entry), out_dir_path)


# the function starts here.
browser()
captcha()
xpath()
print_func()

time.sleep(2)
block_func()
block = block_number.options

for i in range(1, len(block)):  # To download any one block Change here.   :)
    block_number_list.append(block[i].text)

print(block_number_list)
for val in block_number_list:
    while t != 0 and len(block_number_list) >> 1:
        xpath()
        captcha()
        block_func()
        break
    t += 1
    block_number.select_by_value(val)
    time.sleep(2)
    survey_func()
    survey = survey_number.options
    survey_number_list = []
    for i in range(15, len(survey)):  # In case of any error in future change the value here
        survey_number_list.append(survey[i].text)

    for ele in survey_number_list:
        q = 0
        while w != 0 and len(survey_number_list) >> 1:
            xpath()
            block_func()
            block_number.select_by_value(val)
            captcha()
            survey_func()
            break

        w += 1
        survey_number.select_by_value(ele)
        sub_division_list = []

        time.sleep(2)
        sub_division_func()

        sub_division = sub_division_number.options
        for j in range(42, len(sub_division)):
            sub_division_list.append(sub_division[j].text)  # This should run only once.

        for k in sub_division_list:
            while q != 0 and len(sub_division_list) >> 1:
                xpath()
                block_func()
                block_number.select_by_value(val)
                captcha()
                survey_func()
                survey_number.select_by_value(ele)
                sub_division_func()
                break

            q += 1
            sub_division_number.select_by_value(k)
            submit()
            try:
                driver.find_element(By.XPATH, '//*[@id="auto_off"]/div[29]/div/font/strong')
            except NoSuchElementException:
                pass
            else:
                repeat(val, ele, k)
            finally:
                download(ele, k)
                time.sleep(3)
                browser()
    move(val)

end = process_time()
print('The time taken is :', (end - start))
