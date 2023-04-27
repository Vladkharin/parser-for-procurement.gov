from selenium import webdriver
from fake_useragent import UserAgent
import time
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException       
import xlsxwriter 
import pandas as pd
import numpy as np
from selenium.webdriver.chrome.service import Service
import re
import csv
import os

def excel(name_file):
    data = pd.read_excel(f'C:\\Users\\koooo\\Desktop\\для отправки\\{name_file}.xlsx', usecols='B')
    df = pd.DataFrame(data)
    BIN = df.values.tolist()
    list_BIN = []
    for i in BIN:
        for x in i:
            x = str(x)
            x = re.sub("[^0-9]", "", x)
            if len(x) == 9:
                x = '000' + x
                list_BIN.append(x)
            elif len(x) == 10:
                x = '00' + x
                list_BIN.append(x)
            elif len(x) == 11:
                x = '0' + x
                list_BIN.append(x)
            else:
                list_BIN.append(x)
                pass
    list_BIN_sort = list(set(list_BIN))
    return list_BIN_sort

def search_url(name_file):
    try:
        z = 1
        list_object = []
        for item in excel(name_file):
            print(z)
            useragent = UserAgent()
            try:
                options = webdriver.FirefoxOptions()
                options.set_preference("general.useragent.override", useragent.random)
                options.set_preference('dom.webdriver.enabled', False)
                options.headless = True
                service=Service(r"C:\Users\koooo\Desktop\procurement.gov\firefoxdriver\geckodriver.exe")
                driver = webdriver.Firefox(
                    service=service,
                    options=options
                )
                driver.get('https://procurement.gov.kz/ru/registry/supplierreg')
                time.sleep(2)

                bin_input = driver.find_element(By.ID, 'in_name')
                bin_input.clear()
                bin_input.send_keys(item)
                bin_input.send_keys(Keys.ENTER)
                time.sleep(2)

                name_button = driver.find_element(By.CLASS_NAME, 'odd')
                name = name_button.find_element(By.TAG_NAME, 'a')
                name.click()
                time.sleep(2)

                panel = driver.find_elements(By.TAG_NAME, 'table')
                trs = panel[0].find_elements(By.TAG_NAME, 'tr')
                object_text = {
                        'Компания': '',
                        'Регион': '',
                        'БИН': '',
                        'email': '',
                        'Телефон': ''
                    }
                for tr in trs:
                    td = tr.find_element(By.TAG_NAME, 'td')
                    th = tr.find_element(By.TAG_NAME, 'th')
                    try:
                        if th.text == 'Регион':
                            object_text['Регион'] = td.text
                        elif th.text == 'Контактный телефон:':
                            object_text['Телефон'] = td.text
                        elif th.text == 'Контактный телефон:':
                            object_text['Телефон'] = td.text
                        elif th.text == 'БИН участника':
                            object_text['БИН'] = td.text
                        elif th.text == 'E-Mail:':
                            object_text['email'] = td.text
                        elif th.text == 'Наименование на рус. языке':
                            object_text['Компания'] = td.text
                    except:
                        continue
                print(object_text)
                list_object.append(object_text)
            except:
                continue
            finally:
                z += 1
                driver.close()
                driver.quit()
    finally:
        print(list_object)
        return list_object         

def writer(name_file):
    try:
        book = xlsxwriter.Workbook(f'C:\\Users\\koooo\\Desktop\\procurement.gov\\{name_file}\\сортировка {name_file}.xlsx')
        page = book.add_worksheet('')
        row = 0
        column = 0
        page.set_column('A:A', 50)
        page.set_column('B:B', 100)
        page.set_column('C:C', 20)
        page.set_column('D:D', 40)
        page.set_column('E:E', 40)
       
        for item in search_url(name_file):
            print(item)
            page.write(row, column, item['Компания'])
            page.write(row, column+1, item['Регион'])
            page.write(row, column+2, item['БИН'])
            page.write(row, column+3, item['email'])
            page.write(row, column+4, item['Телефон'])
            row += 1
    finally:
        book.close()

def sort_telephone_list(name_file):
    data_phone = pd.read_excel(f'C:\\Users\\koooo\\Desktop\\procurement.gov\\{name_file}\\сортировка {name_file}.xlsx', usecols='E')
    df_phone = pd.DataFrame(data_phone)
    telephone = df_phone.values.tolist()
    new_list_telephone = np.array(telephone).flatten()

    data_name = pd.read_excel(f'C:\\Users\\koooo\\Desktop\\procurement.gov\\{name_file}\\сортировка {name_file}.xlsx', usecols='A')
    df_name = pd.DataFrame(data_name)
    name = df_name.values.tolist()
    new_list_name = np.array(name).flatten()

    data_bin = pd.read_excel(f'C:\\Users\\koooo\\Desktop\\procurement.gov\\{name_file}\\сортировка {name_file}.xlsx', usecols='C')
    df_bin = pd.DataFrame(data_bin)
    any_bin = df_bin.values.tolist()
    new_list_bin = np.array(any_bin).flatten()

    data_email = pd.read_excel(f'C:\\Users\\koooo\\Desktop\\procurement.gov\\{name_file}\\сортировка {name_file}.xlsx', usecols='D')
    df_email = pd.DataFrame(data_email)
    any_email = df_email.values.tolist()
    new_list_email = np.array(any_email).flatten()

    object_lists_csv = []
    for car in range(0, len(telephone)):
        bin_rebuild = str(new_list_bin[car])
        object_csv = {
            "Company" : new_list_name[car],
            "Заметки": bin_rebuild,
            'Primary Phone' : new_list_telephone[car],
            'Email': new_list_email[car]
        }

        # print(object_csv['Email'])

        if object_csv['Primary Phone'] != None and object_csv['Primary Phone'] != "nan":
            object_lists_csv.append(object_csv)
    new_object_lists_cvs = []
    for object_csv in object_lists_csv:
        if type(object_csv['Primary Phone']) != float:
            string_telephones = str(object_csv['Primary Phone'])
            list_telephones = string_telephones.split(' ')
            list_name = object_csv["Company"].split('"')
            bin_build = object_csv['Заметки']
            if object_csv['Email'] == 'nan':
                object_csv['Email'] = ''
            email = object_csv['Email'].split(',')[0]
            str_name = ''
            if list_name[0] == 'Товарищество с ограниченной ответственностью ':
                list_name[0] = 'ТОО '
                str_name = ''.join(list_name)
            for telephone_not_sort in list_telephones:
                telephone = re.sub('[^0-9]', '', telephone_not_sort)
                if telephone != '':
                    if telephone[0] == '8':
                        # print('8')
                        telephone = telephone.replace('8', '+7', 1)
                    elif telephone[0] == '7':
                        # print('7')
                        telephone = telephone.replace('7', '+7', 1)
                    if len(telephone) > 12:
                        telephone = telephone[:11]
                    elif len(telephone) != 12:
                        continue
                if telephone.find('7700') == -1 and telephone.find('7701') == -1 and telephone.find('7702') == -1 and telephone.find('7703') == -1 and telephone.find('7704') == -1 and telephone.find('7705') == -1 and telephone.find('7706') == -1 and telephone.find('7707') == -1 and telephone.find('7708') == -1 and telephone.find('7709') == -1 and telephone.find('7747') == -1 and telephone.find('7750') == -1 and telephone.find('7751') == -1 and telephone.find('7760') == -1 and telephone.find('7761') == -1 and telephone.find('7762') == -1 and telephone.find('7763') == -1 and telephone.find('7764') == -1 and telephone.find('7771') == -1 and telephone.find('7775') == -1 and telephone.find('7776') == -1 and telephone.find('7777') == -1 and telephone.find('7778') == -1:
                    pass
                else:
                    new_object_cvs = {
                        "Company": str_name,
                        "Заметки": bin_build,
                        'Primary Phone': telephone,
                        'Email': email
                    }
                    new_object_lists_cvs.append(new_object_cvs)
    return new_object_lists_cvs

# sort_telephone_list('г.Алмата сортировка')

def createTelFile(name_file):
    try:
        book = xlsxwriter.Workbook(f'C:\\Users\\koooo\\Desktop\\procurement.gov\\{name_file}\\tel {name_file}.xlsx')
        page = book.add_worksheet('')
        row = 1
        column = 0
        page.set_column('A:A', 50)
        for item in sort_telephone_list(name_file):
            page.write(row, column, item['Primary Phone'])
            row += 1
        
    except Exception as ex:
        print(ex)
    finally:
        book.close()

def csv_d(name_file):
    with open(f"{name_file}\\csv {name_file}.csv", mode="w", encoding='utf-8', newline='') as w_file:
        fieldnames = ["Company", 'Primary Phone', "Заметки", 'Email']
        file_writer = csv.DictWriter(w_file, fieldnames=fieldnames)
        file_writer.writeheader()
        for object_csv in sort_telephone_list(name_file):
            file_writer.writerow(object_csv)

def createEmailFile(name_file):
    try:
        book = xlsxwriter.Workbook(f'C:\\Users\\koooo\\Desktop\\procurement.gov\\{name_file}\\email {name_file}.xlsx')
        page = book.add_worksheet('')
        row = 1
        column = 0
        page.set_column('A:A', 50)
        for item in sort_telephone_list(name_file):
            page.write(row, column, item['Email'])
            row += 1
    except Exception as ex:
        print(ex)
    finally:
        book.close()


def all_function_start(dirs):
    # item_list = ['Нурсултан_сортировка']
    # for item in item_list:
    #     # os.mkdir(f'{item}')
    #     # writer(item)
    #     # second_writer(item)
    #     # csv_d(item)
    for dir in dirs:
        list_dir = os.listdir(f'C:\\Users\\koooo\\Desktop\\{dir}')
        # for item in list_dir:
        item = list_dir[0][:-5]
        print(item)
        os.mkdir(f'{item}')
        writer(item)
        print('Writer success')
        createTelFile(item)
        print('Second Writer success')
        csv_d(item)
        print('Csv success')
        createEmailFile(item)
        print('creating Email file sucesses')


all_function_start(['для отправки'])

