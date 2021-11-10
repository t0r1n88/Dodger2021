import pandas as pd
import os
from docxtpl import DocxTemplate
import csv
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
from tkinter import ttk

path ='data/'
# Считываем csv файл, не забывая что екселввский csv разделен на самомо деле не запятыми а точкой с запятой
reader = csv.DictReader(open('data/data.csv'), delimiter=';')
# Конвертируем объект reader в список словарей
data = list(reader)
# Получаем значения колонок документа с данными

for row in data:
    doc = DocxTemplate('data/Template.docx')
    context = row
    # Превращаем строку в список кортежей, где первый элемент кортежа это ключ а второй данные
    id_row = list(row.items())
    doc.render(context)
    doc.save(f'{id_row[0][1]}.docx')
