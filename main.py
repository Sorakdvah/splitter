import math
import os
import tkinter as tk
from tkinter import filedialog, simpledialog
import pandas as pd


def split_excel_file(file_name, number_of_parts):
    # Чтение файла Excel в DataFrame
    df = pd.read_excel(file_name, engine='openpyxl')

    # Определение наименьшего количества строк в каждом фрагменте
    rows_per_part = math.ceil(df.shape[0] / number_of_parts)

    # Создание и сохранение фрагментов Excel
    for i in range(number_of_parts):
        # Выборка строк для текущего фрагмента
        start_row = i * rows_per_part
        end_row = min((i + 1) * rows_per_part, df.shape[0])
        part_df = df[start_row:end_row]

        # Создание имени файла и пути для сохранения файла
        file_basename, file_extension = os.path.splitext(file_name)
        output_file_name = f"{file_basename}_{i + 1}{file_extension}"

        # Сохранение фрагмента в файл Excel
        part_df.to_excel(output_file_name, index=False, engine='openpyxl')

    print(f"Файл {file_name} разбит на {number_of_parts} частей.")

root = tk.Tk()
root.withdraw()

while True:
    # Открытие диалогового окна выбора файла
    file_name = filedialog.askopenfilename(title="Выберите файл Excel", filetypes=[("Файлы Excel", "*.xlsx"), ("Файлы Excel", "*.xls")])

    if file_name:
        # Отображение окна для ввода количества частей
        number_of_parts = simpledialog.askinteger("Разделение файла", "Введите количество частей:")

        if number_of_parts and number_of_parts > 0:
            split_excel_file(file_name, number_of_parts)
        else:
            print("Некорректное количество частей.")
    else:
        break