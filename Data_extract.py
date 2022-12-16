from tkinter import ttk
import tkinter as tk
from tkinter import *
import os

import time
import pandas as pd
import threading
import chromedriver_autoinstaller
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from list import *

try:
    from ctypes import windll

    windll.shcore.SetProcessDpiAwareness(1)
except:
    pass

root = tk.Tk()
root.title("Village Excel Generator")
chromedriver_autoinstaller.install()

app_width = 800
app_height = 600

screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()

x = (screen_width / 2) - (app_width / 2)
y = (screen_height / 2) - (app_height / 2)

root.geometry(f'{app_width}x{app_height}+{int(x)}+{int(y)}')

# Defining the global variables
block_number_list = []
district_list = []
k = 0
survey_number, sub_division_number, block_number = 1, 1, 1
block, district, town, taluk, ward = 0, 0, 0, 0, 0
driver = 0
district_value = tk.StringVar()
taluk_value = tk.StringVar()
town_value = tk.StringVar()
ward_value = tk.StringVar()
block_from_value = tk.StringVar()
block_to_value = tk.StringVar()


def exe_script(*args):
    threading.Thread(target=start_func).start()


def close_func(*args):
    threading.Thread(target=root.destroy).start()


def taluk_func(*args):
    taluk_list = []
    for p, q in values.items():
        if district_value.get() == p:
            for m, n in q.items():
                taluk_list.append(m)
    taluk_input.config(value=taluk_list)


def town_func(*args):
    town_list = []
    for p, q in values.items():
        for m, n in q.items():
            if taluk_value.get() == m:
                for a, b in n.items():
                    town_list.append(a)
    town_input.config(value=town_list)


def ward_func(*args):
    ward_list = []
    for p, q in values.items():
        for m, n in q.items():
            for a, b in n.items():
                if town_value.get() == a:
                    ward_list = b
    ward_input.config(value=ward_list)


def browser():
    driver.get('https://eservices.tn.gov.in/eservicesnew/land/chittaCheckNewUrban_en.html?lan=en')
    driver.maximize_window()


def xpath():
    global district, town, taluk, ward
    district = Select(driver.find_element(By.XPATH, '//*[@id="districtCode"]'))
    taluk = Select(driver.find_element(By.XPATH, '//*[@id="auto_off"]/div[8]/div[2]/select'))
    town = Select(driver.find_element(By.XPATH, '//*[@id="auto_off"]/div[11]/div[2]/select'))
    ward = Select(driver.find_element(By.XPATH, '//*[@id="auto_off"]/div[14]/div[2]/select'))

    district.select_by_visible_text(district_input.get())
    driver.implicitly_wait(10)
    taluk.select_by_visible_text(taluk_input.get())
    town.select_by_visible_text(town_input.get())
    ward.select_by_visible_text(ward_input.get())


def check_file(element):
    error_list = ['?', ':', '/']
    for i in error_list:
        element = element.replace(i, '-')
    return element


def block_func():
    global block_number
    block_number = Select(driver.find_element(By.XPATH, '//*[@id="auto_off"]/div[15]/div[2]/select'))


def survey_func():
    global survey_number
    survey_number = Select(driver.find_element(By.NAME, 'surveyNo'))


def sub_division_func():
    global sub_division_number
    sub_division_number = Select(driver.find_element(By.NAME, 'subdivNo'))


def start_func(*args):
    global block_to_value, block_from_value, block, driver
    driver = webdriver.Chrome()
    browser()
    xpath()
    time.sleep(2)
    block_func()
    block = block_number.options

    if not block_from_input.get() == '':
        block_from_value = int(block_from_input.get())
    else:
        block_from_value = int(1)

    if not block_to_input.get() == '':
        block_to_value = int(block_to_input.get())
    else:
        block_to_value = len(block)

    for i in range(block_from_value, block_to_value):  # To download any one block Change here.   :)
        block_number_list.append(block[i].text)
    t = 0
    path = r"E:\Verification_Files\Town-" + check_file(town_input.get()) + " Ward-" + check_file(ward_input.get())

    if not os.path.exists(path):  # create folder if it doesn't exist
        os.makedirs(path)
    messagebox.showerror("Error", "Error message")
    for val in block_number_list:
        total = 0
        templist = []
        while t != 0 and len(block_number_list) >> 1:
            driver.get('https://eservices.tn.gov.in/eservicesnew/land/chittaCheckNewUrban_en.html?lan=en')
            xpath()
            block_func()
            break
        t += 1
        block_number.select_by_value(val)
        time.sleep(2)
        survey_func()
        survey = survey_number.options
        survey_number_list = []
        for i in range(1, len(survey)):  # In case of any error in future change the value here
            survey_number_list.append(survey[i].text)
        w = 0
        for ele in survey_number_list:

            while w != 0 and len(survey_number_list) >> 1:
                xpath()
                block_func()
                block_number.select_by_value(val)
                survey_func()
                break

            w += 1
            survey_number.select_by_value(ele)
            sub_division_list = []

            time.sleep(2)
            sub_division_func()

            sub_division = sub_division_number.options

            for j in range(1, len(sub_division)):
                sub_division_list.append(sub_division[j].text)
                total += 1  # This should run only once.

            Table_dict = {'Survey Number': ele,
                          'Sub-Division Number': sub_division_list,
                          'Total': total}
            templist.append(Table_dict)
            df = pd.DataFrame(templist)

        file = r"D:\Verification_Files\Town-" + check_file(town_input.get()) + " Ward-" + check_file(ward_input.get()) \
               + r"\Block-" + val + ".xlsx"
        df.to_excel(file)


for key, value in values.items():
    district_list.append(key)

district_label = Label(root, text="District")
district_label.grid(row=0, column=0, padx=(40, 10), pady=(50, 10))
district_input = ttk.Combobox(root, values=district_list, textvariable=district_value, width=25, state="readonly")
district_input.grid(row=0, column=1, padx=(40, 10), pady=(50, 10))
district_input.set(value="Please Select...")
district_input.bind("<<ComboboxSelected>>", taluk_func)

taluk_label = Label(root, text="Taluk")
taluk_label.grid(row=1, column=0, padx=(40, 10), pady=(10, 10))
taluk_input = ttk.Combobox(root, textvariable=taluk_value, width=25, state="readonly")
taluk_input.grid(row=1, column=1, padx=(40, 10), pady=(10, 10))
taluk_input.set(value="Please Select...")
taluk_input.bind("<<ComboboxSelected>>", town_func)

town_label = Label(root, text="Town")
town_label.grid(row=2, column=0, padx=(40, 10), pady=(10, 10))
town_input = ttk.Combobox(root, textvariable=town_value, width=25, state="readonly")
town_input.grid(row=2, column=1, padx=(40, 10), pady=(10, 10))
town_input.set(value="Please Select...")
town_input.bind("<<ComboboxSelected>>", ward_func)

ward_label = Label(root, text="Ward")
ward_label.grid(row=3, column=0, padx=(40, 10), pady=(10, 10))
ward_input = ttk.Combobox(root, textvariable=ward_value, width=25, state="readonly")
ward_input.grid(row=3, column=1, padx=(40, 10), pady=(10, 10))
ward_input.set(value="Please Select...")

block_label = Label(root, text="Block")
block_label.grid(row=4, column=0, padx=(40, 10), pady=(20, 10))
block_from = Label(root, text="From")
block_from.grid(row=5, column=0, padx=(40, 10), pady=(5, 10))
block_from_input = Entry(root, textvariable=block_from_value)
block_from_input.grid(row=5, column=1, padx=(40, 10), pady=(5, 10))
block_to = Label(root, text="To")
block_to.grid(row=5, column=2, padx=(40, 10), pady=(5, 10))
block_to_input = Entry(root, textvariable=block_to_value)
block_to_input.grid(row=5, column=3, padx=(40, 10), pady=(5, 10))

submit = Button(root, text='Generate', command=exe_script)
submit.grid(row=6, column=1, padx=(40, 40), pady=(50, 20))
Quit = Button(root, text='Quit', command=close_func)
Quit.grid(row=6, column=2, padx=(40, 40), pady=(50, 20))
root.bind("<Return>", exe_script)
root.bind("<<KP_Enter>>", exe_script)
root.mainloop()
