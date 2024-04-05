import pandas as pd
import os
import math
import tkinter as tk
from tkinter import filedialog

#Функция округления
def custom_ceil(number):
    if number < 0:
        return math.floor(number)
    else:
        return math.ceil(number)


# Функция для обработки пива
def handle_beer(nomenclature, stock, total_quantity):
    forecast = math.floor((stock - total_quantity) / 30)
    if (forecast >= 0):
        forecast = 0
        remaining_liters = stock - total_quantity
    else:
        remaining_liters = 0
    return {'Номенклатура': nomenclature, 'Остаток': stock, 'Прогноз': total_quantity, 'Заказ кег': forecast, 'Остаток литров': remaining_liters}


def load_sales_file():
    global sales_file_path
    sales_file_path = filedialog.askopenfilename()
    print(f"Выбран файл с продажами: {sales_file_path}")
    if sales_file_path:
        message_label_sales.config(text=f"{os.path.basename(sales_file_path)}")
    else:
        message_label_sales.config(text="Файл не выбран")

def load_residuals_file():
    global residuals_file_path
    residuals_file_path = filedialog.askopenfilename()
    print(f"Выбран файл с остатками: {residuals_file_path}")
    if residuals_file_path:
        message_label_residuals.config(text=f"{os.path.basename(residuals_file_path)}")
    else:
        message_label_residuals.config(text="Файл не выбран")


def generate_report():

    if sales_file_path is None or residuals_file_path is None:
        print("Пожалуйста, выберите оба файла перед генерацией отчета.")
        return

    # Чтение данных из первого файла (с продажами)
    sales = pd.read_excel(f'{sales_file_path}', skiprows=4, usecols=[0, 3, 4, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17])

    sales_column_names = ['Магазин', 'Дата и время', 'Id чека', 'Диск карт', 'Владелец карты', 'Номер телефона', 'Категория товара', 'Группа товара', 'Номенклатура', 'Сумма продаж', 'Количество товара', 'Остаток на складе', 'Сумма скидки', 'Себестоимость продаж', 'Валовая прибыль']
    sales.columns = sales_column_names

    print(sales.columns)


    # Чтение данных из второго файла (с остатками)
    residuals = pd.read_excel(f'{residuals_file_path}', skiprows=2, usecols=[0, 2]) # пропуск строки с назва
    residuals_column_names = ['Номенклатура', 'Остаток']
    residuals.columns = residuals_column_names

    print(residuals.columns)

    # Создание пустого DataFrame для результатов
    results = pd.DataFrame(columns=['Номенклатура', 'Остаток', 'Прогноз', 'Заказ кег', 'Остаток литров'])

    # Обработка данных
    for index, row in residuals.iterrows():
        nomenclature = row['Номенклатура']
        stock = row['Остаток']

        # Находим все записи в первом файле, соответствующие текущей номенклатуре
        matching_rows = sales[sales['Номенклатура'] == nomenclature]

        if not matching_rows.empty and matching_rows['Категория товара'].str.contains('Пиво').any():
            total_quantity = matching_rows['Количество товара'].sum()

            result = handle_beer(nomenclature, stock, total_quantity)

            # Добавление результатов в DataFrame
            results = results._append(result, ignore_index=True)

            # Замена именований
            results['Номенклатура'] = results['Номенклатура'].replace("Амбирлэнд Жигулевское", "Жигулевское")
            results['Номенклатура'] = results['Номенклатура'].replace("Амбирлэнд Пилснер нефильтрованное", "Пилснер н/ф")
            results['Номенклатура'] = results['Номенклатура'].replace("Амбирлэнд Пилснер фильтрованное", "Пилснер ф")
            results['Номенклатура'] = results['Номенклатура'].replace("Амбирлэнд Пшеничное", "Вайс")
            results['Номенклатура'] = results['Номенклатура'].replace("Амбирлэнд Светлое нефильтрованное", "Боровское н/ф")
            results['Номенклатура'] = results['Номенклатура'].replace("Амбирлэнд Светлое фильтрованное", "Бундес")
            results['Номенклатура'] = results['Номенклатура'].replace("Амбирлэнд Темное", "Темное")
            results['Номенклатура'] = results['Номенклатура'].replace("Амбирлэнд Вишневый крик", "Вайлд Черри")


    # Запись результатов в новый Excel файл
    if os.path.exists('Прогноз_пиво_короткое_время.xlsx'):
        # Если файл существует, удаляем его
        os.remove('Прогноз_пиво_короткое_время.xlsx')

    second_table = pd.DataFrame(columns=['Номенклатура', 'Заказ кег'])

    for index, row in results.iterrows():
        nomenclature = row['Номенклатура']
        forecast = row['Заказ кег']
        if nomenclature == "Вайлд Черри":
            second_table = second_table._append({'Номенклатура': nomenclature, 'Заказ кег': f"{abs(forecast)}*20"}, ignore_index=True)
        else:
            second_table = second_table._append({'Номенклатура': nomenclature, 'Заказ кег': f"{abs(forecast)}*30"}, ignore_index=True)

    # Запись двух таблиц в один Excel файл
    with pd.ExcelWriter('Прогноз_пиво_короткое_время.xlsx') as writer:
        results.to_excel(writer, index=False)
        second_table.to_excel(writer, startrow=len(results) + 3, index=False)
        print("Генерация отчета...")
        window.destroy()



window = tk.Tk()
window.title("Приложение")
window.geometry("500x150")
message_label_sales = tk.Label(window, text="Файл не выбран")
message_label_sales.place(x=65, y=40)
message_label_residuals = tk.Label(window, text="Файл не выбран")
message_label_residuals.place(x=340, y=40)


btn_sales = tk.Button(window, text="Загрузить файл с продажами", command=load_sales_file)
btn_residuals = tk.Button(window, text="Загрузить файл с остатками", command=load_residuals_file)
btn_sales.place(x=25, y=10)
btn_residuals.place(x=300, y=10)
btn_generate = tk.Button(window, text="Сгенерировать отчет", command=generate_report)
btn_generate.place(x=185, y=100)
window.mainloop()

print("Все успешно!")
