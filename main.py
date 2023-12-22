# -*- coding: utf-8 -*-
from bs4 import BeautifulSoup
import requests
import json
import os
from openpyxl import Workbook
import csv

exel_list = []

# Получить путь к рабочему столу
current_directory = os.path.join(os.path.expanduser('~'), 'Desktop')
# Получить путь к директории, в которой находится файл main.py
#current_directory = os.path.dirname(os.path.abspath(__file__))

# Сформировать полный путь к файлу txt
file_path = os.path.join(current_directory, "parse.txt")

def parse_json_script(soup):

    exel_list.append(url)

    if soup.findAll("title"):
        title = soup.find("title").string
        #print("Title: "+soup.find("title").string)
        exel_list.append(title)
    else:
        exel_list.append("Title не взят")

    if soup.findAll("h1"):
        h1 = soup.find("h1").string
        #print("h1: "+soup.find("h1").string)
        exel_list.append(h1)
    else:
        exel_list.append("h1 не взят")

    if soup.findAll("meta", attrs={"name": "description"}):
        description = soup.find("meta", attrs={"name": "description"}).get("content")
        #print("Description: "+soup.find("meta", attrs={"name": "description"}).get("content"))
        exel_list.append(description)
    else:
        exel_list.append("Description не взят")

    return exel_list

def export_to_excel(data, filename):
    # Сохранить файл
    excel_filepath = os.path.join(current_directory, filename)

    # Создаем новую книгу Excel
    workbook = Workbook()
    sheet = workbook.active

   # Заполняем лист данными из многомерного массива (если урлов много)
    for row_index, row in enumerate(data, start=1):
        for col_index, value in enumerate(row, start=1):
            sheet.cell(row=row_index, column=col_index, value=value)

    # Сохраняем книгу Excel
    workbook.save(excel_filepath)
    print(f"Данные успешно экспортированы в файл: {excel_filepath}")


def export_to_csv(data, filename):
    # Сохранить файл
    file_path = os.path.join(current_directory, filename)

    # Открываем файл для записи
    with open(file_path, 'w', newline='', encoding='utf-8') as csv_file:
        # Создаем объект writer для записи в CSV файл
        csv_writer = csv.writer(csv_file)

        # Записываем данные из массива в CSV файл
        csv_writer.writerows(data)

    print(f'Данные успешно экспортированы в {file_path}')


with open(file_path, 'r') as file:
    unique_lines = set()  # Используем множество для хранения уникальных строк
    duplicate_lines = set()  # Используем множество для хранения дублирующихся строк
    all_lines = []  # Используем список для хранения всех строк

    urls_list = file.readlines()

    # Проверка каждой строки на уникальность и наличие текста
    for line in urls_list:
        cleaned_line = line.strip()  # Удаление символов новой строки и пробелов
        if cleaned_line:
            all_lines.append(cleaned_line)

            if cleaned_line in unique_lines:
                duplicate_lines.add(cleaned_line)
            else:
                unique_lines.add(cleaned_line)

    # Вывод количества дублирующихся и уникальных строк
    print(f'Количество удаленных дублирующихся строк: {len(duplicate_lines)}')
    print(f'Количество уникальных строк: {len(unique_lines)}')

    # Преобразование множества уникальных строк обратно в массив
    urls_list = list(unique_lines)

    # Количество строк в файле
    lines = urls_list

    # Удаление символов новой строки из каждой строки
    urls_list = [line.strip() for line in urls_list]

#Массив массивов для экспорта
data = []

url_count = 0

try:
    for url in urls_list:
        headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'}
        response = requests.get(url, headers=headers)
        if(response.status_code == 200): soup = BeautifulSoup(response.content, 'html.parser')

        parse_json_script(soup)
        print(url)
        print(exel_list)
        url_count = url_count + 1
        print(str(url_count), "из", len(lines), " строк обработано")

        # Добавляем каждый отдельный массив данных по урлу в единый массив, который будет затем экспортироваться
        data.append(exel_list)
        exel_list = []
except ConnectionRefusedError:
    print("Ошибка: Сервер отклонил соединение.")

except Exception as e:
    print(f"Произошла ошибка: {e}")

finally:
    export_to_excel(data, "output_metadata.xlsx")

#print(exel_list)

#print(data)
# Экспорт данных в файл
#export_to_excel(data, "output_metadata.xlsx")
#export_to_csv(exel_list, "output.csv")
