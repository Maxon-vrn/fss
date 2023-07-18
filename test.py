import time
import os
from numpy import block
import requests
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver import chrome
from selenium.webdriver.chrome.service import Service
import pandas as pd



print()
def zapros(link):
    """Проверяет доступность для работы сраницы и выводит  responce"""
    header = {'User-Agent':"Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/102.0.5005.115 Safari/537.36 OPR/88.0.4412.74"}
    sait = requests.get(link, headers=header)
    return sait

def find_file_in_folder() :
    """Исчет в текущей директории файлы с расширениеом .ods и сохраняет в список"""
    path = './' #путь в текущей директории
    try:
        directory = os.listdir(path)
        list_file = []
        for i in directory: #в списке из файлов в текущей папке исчет нужное расширение
            if i.endswith('.ods'):
                list_file.append(i)
        return list_file
    except Exception as e:
        print(f"Ненайдены файлы с расширением '.ods' в текущей директории. Ошибка ->  {e}")


def open_file(path):
    """Обработчик файла Exel который достает данные и сохраняет в словарь инн для обработки """

    # Load the xlsx file
    #usecols='F'-столбец,header=6- строка+1,nrows=3- колонка+1,sheet_name=0 -лист чтения
    excel_data = pd.read_excel(path, usecols='F',header=5,sheet_name=0)
    # не обрабатывает xls и тп pd.read_excel(io='temp1.xlsx', engine='openpyxl')

    dict_inn = excel_data.to_dict() #преобразование данных в словарь для обработки данных
    return (dict_inn)



def saver_file(text):   #сохраняем данные в новый файл
    """Сохраняет полученный результат в файл/таблицу"""

    try:
        with open('text.txt', 'a') as f:
            f.write(text + '\n')
        print('sucsess')
    except Exception as e:
        print(f'Неудалось сохранить текст в файл, причина: {e}')



def parsing(list_file):
    """ """
    #for i in list_file:
    #   print(i) - путь к файлу в текущей директории
    try:
        dict_inn = open_file(list_file[1])  # list_file[0] - юр лицо ,list_file[0] - ипешники
    except Exception as e:
        print(f'Неудалось получить данные из файла Exel, причина: {e}')

    try:

        driver = webdriver.Chrome()  # Открываем Firefox
        driver.maximize_window()

        for i in dict_inn.values(): #итерация вложенного словаря для доступа к инн
            for id, inn in i.items():  #inn,id
                url = f'https://fedresurs.ru/search/entity?code={inn}&regionNumber=%D0%92%D1%81%D0%B5&onlyActive=true'

                # открываем страницу с уже вставленным значением инн
                driver.get(url)  # Открываем страницу
                time.sleep(2)

                # парсим страницу для уточнения статуса по инн
                soup = BeautifulSoup(driver.page_source, 'html.parser')  # Получаем готовый html и парсим его
                td_list = soup.find_all('td', class_='with_link')  # Получаем список
                # print(td_list)

                try:
                    p_value = td_list[0].find_all('p')  # Получаем из списока значение/результат
                    # print(p_value)
                    for i in p_value:
                        if i.text == 'н/д ':
                            information_to_save = f'{id+1} по данному ИНН {inn} -> нет данных'
                            saver_file(information_to_save)    #запиысываемы нужные данные в файл
                            print(f'{id+1} по данному ИНН {inn} -> нет данных')
                        else:
                            print(f'{id+1} по данному ИНН {inn} -> {i.text}')
                except Exception as e:
                    print(f'{id+1} по данному ИНН {inn} -> "ненайдено" ->{e}')  # list index out of range"""


    except Exception as e:
        print(f'Ошибка проверки/обработки данных на сайте: {e}')



link = 'https://fedresurs.ru/search/entity?regionNumber=%D0%92%D1%81%D0%B5&onlyActive=true'
try:
    if str(zapros(link)) == '<Response [200]>':
        print('Соединение установленo')
        parsing(find_file_in_folder())


except Exception as e:
    print(f'Не удалось подключиться к сайту: {e}')