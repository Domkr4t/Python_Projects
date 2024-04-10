import pandas as pd
import os
import math
import tkinter as tk
from tkinter import filedialog
import platform


sales_file_path = None
residuals_file_path = None
report_file_path = None


#Функция округления
def custom_ceil(number):
    if number - int(number) >= 0.15:
        return math.ceil(number)
    else:
        return math.floor(number)


# Функция для обработки пива
def handle_beer(nomenclature, stock, total_quantity):
    forecast = math.floor((stock - total_quantity) / 30)
    if (forecast >= 0):
        forecast = 0
        remaining_liters = stock - total_quantity
    else:
        remaining_liters = 0
    return {'Номенклатура': nomenclature, 'Остаток': stock, 'Прогноз': total_quantity, 'Заказ кег': forecast, 'Остаток литров': remaining_liters}


# Функция для обработки закусок к пиву
def handle_snacks(nomenclature, stock, total_quantity):
    forecasted_balance = stock - total_quantity
    if (forecasted_balance < 600 and not forecasted_balance.is_integer()):
        forecast = 1 - forecasted_balance
    elif (forecasted_balance >= 600 and not forecasted_balance.is_integer()):
        forecast = 0
    elif (forecasted_balance.is_integer()):
        forecast = 0

    return {'Номенклатура': nomenclature, 'Остаток': stock, 'Прогноз': total_quantity, 'Прогнозируемый остаток': forecasted_balance, 'Заказ': abs(forecast)}


# Функция для обработки прочего
def handle_other(nomenclature, stock, total_quantity):
    forecasted_balance = stock - total_quantity
    if (forecasted_balance < 600 and not forecasted_balance.is_integer()):
        forecast = 1 - forecasted_balance
    elif (forecasted_balance >= 600 and not forecasted_balance.is_integer()):
        forecast = 0
    elif (forecasted_balance.is_integer()):
        forecast = 0

    return {'Номенклатура': nomenclature, 'Остаток': stock, 'Прогноз': total_quantity, 'Прогнозируемый остаток': forecasted_balance, 'Заказ': abs(forecast)}


# Загрузка файла с продажами
def load_sales_file():
    global sales_file_path
    sales_file_path = filedialog.askopenfilename(initialdir='c:/User/Desktop', title='Загрузка файла с продажами')
    print(f"Выбран файл с продажами: {sales_file_path}")
    if sales_file_path:
        message_label_sales.config(text=f"{os.path.basename(sales_file_path)}", foreground='green')
        message_file_not_upload.config(text=" ")
    else:
        message_label_sales.config(text="Файл не выбран", foreground='red')
        sales_file_path = None


# Загрузка файла с остатками
def load_residuals_file():
    global residuals_file_path
    residuals_file_path = filedialog.askopenfilename(initialdir='c:/User/Desktop', title='Загрузка файла с остатками')
    print(f"Выбран файл с остатками: {residuals_file_path}")
    if residuals_file_path:
        message_label_residuals.config(text=f"{os.path.basename(residuals_file_path)}", foreground='green')
        message_file_not_upload.config(text=" ")
    else:
        message_label_residuals.config(text="Файл не выбран", foreground='red')
        residuals_file_path = None


# Открытие файла
def open_file(file_path):
    # Открывает файл с помощью системной команды.
    if platform.system() == "Windows":
        os.startfile(file_path)
    elif platform.system() == "Darwin":
        os.system(f"open {file_path}")
    else:
        os.system(f"xdg-open {file_path}")


# Сохранение файла
def save_file():
    global report_file_path
    # Открываем диалог выбора файла для сохранения
    report_file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", title='Сохранить отчет', filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")])


# Создание отчета
def generate_report():

    # Проверка на заполнение файлов
    if sales_file_path is None and residuals_file_path is None:
        message_file_not_upload.config(text="Файлы не загружены")
        return
    elif sales_file_path is None:
        message_file_not_upload.config(text="Продажи не загружены")
        return
    elif residuals_file_path is None:
        message_file_not_upload.config(text="Остатки не загружены")
        return

    # Чтение данных из первого файла (с продажами)
    sales = pd.read_excel(f'{sales_file_path}', skiprows=4, usecols=[0, 3, 4, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17])
    sales_column_names = ['Магазин', 'Дата и время', 'Id чека', 'Диск карт', 'Владелец карты', 'Номер телефона', 'Категория товара', 'Группа товара', 'Номенклатура', 'Сумма продаж', 'Количество товара', 'Остаток на складе', 'Сумма скидки', 'Себестоимость продаж', 'Валовая прибыль']
    sales.columns = sales_column_names

    # Чтение данных из второго файла (с остатками)
    residuals = pd.read_excel(f'{residuals_file_path}', skiprows=2, usecols=[0, 2]) # пропуск строки с назва
    residuals_column_names = ['Номенклатура', 'Остаток']
    residuals.columns = residuals_column_names

    # Создание "шапок" для таблиц
    beer_results = pd.DataFrame(columns=['Номенклатура', 'Остаток', 'Прогноз', 'Заказ кег', 'Остаток литров'])
    beer_second_table = pd.DataFrame(columns=['Номенклатура', 'Заказ кег'])
    snacks_results = pd.DataFrame(columns=['Номенклатура', 'Остаток', 'Прогноз', 'Прогнозируемый остаток', 'Заказ'])
    snacks_second_table = pd.DataFrame(columns=['Номенклатура', 'Заказ'])
    other_results = pd.DataFrame(columns=['Номенклатура', 'Остаток', 'Прогноз', 'Прогнозируемый остаток', 'Заказ'])
    other_second_table = pd.DataFrame(columns=['Номенклатура', 'Заказ'])


    # Обработка данных
    for index, row in residuals.iterrows():

        nomenclature = row['Номенклатура']
        stock = row['Остаток']

        # Находим все записи в первом файле, соответствующие текущей номенклатуре
        matching_rows = sales[sales['Номенклатура'] == nomenclature]

        # Обработка данных для пива
        if not matching_rows.empty and matching_rows['Категория товара'].str.contains("Пиво").any():
            total_quantity = matching_rows['Количество товара'].sum()

            beer_result = handle_beer(nomenclature, stock, total_quantity)
            beer_results = beer_results._append(beer_result, ignore_index=True)
            # Замена именований
            beer_results['Номенклатура'] = beer_results['Номенклатура'].replace("Амбирлэнд Жигулевское", "Жигулевское")
            beer_results['Номенклатура'] = beer_results['Номенклатура'].replace("Амбирлэнд Пилснер нефильтрованное", "Пилснер н/ф")
            beer_results['Номенклатура'] = beer_results['Номенклатура'].replace("Амбирлэнд Пилснер фильтрованное", "Пилснер ф")
            beer_results['Номенклатура'] = beer_results['Номенклатура'].replace("Амбирлэнд Пшеничное", "Вайс")
            beer_results['Номенклатура'] = beer_results['Номенклатура'].replace("Амбирлэнд Светлое нефильтрованное", "Боровское н/ф")
            beer_results['Номенклатура'] = beer_results['Номенклатура'].replace("Амбирлэнд Светлое фильтрованное", "Бундес")
            beer_results['Номенклатура'] = beer_results['Номенклатура'].replace("Амбирлэнд Темное", "Темное")
            beer_results['Номенклатура'] = beer_results['Номенклатура'].replace("Амбирлэнд Вишневый крик", "Вайлд Черри")
        # Обработка данных для закусок к пиву
        if not matching_rows.empty and matching_rows['Категория товара'].str.contains("Закуски к пиву").any() | matching_rows['Категория товара'].str.contains("Рыба").any():
            total_quantity = matching_rows['Количество товара'].sum()

            snacks_result = handle_snacks(nomenclature, stock, total_quantity)
            snacks_results = snacks_results._append(snacks_result, ignore_index=True)

        if not matching_rows.empty and matching_rows['Категория товара'].str.contains("Прочее").any():
            total_quantity = matching_rows['Количество товара'].sum()

            other_result = handle_other(nomenclature, stock, total_quantity)
            other_results = other_results._append(other_result, ignore_index=True)

    #Вторая таблица для пива
    for index, row in beer_results.iterrows():
        nomenclature = row['Номенклатура']
        forecast = row['Заказ кег']

        if nomenclature == "Вайлд Черри":
            beer_second_table = beer_second_table._append({beer_second_table.columns[0]: nomenclature, beer_second_table.columns[1]: f"{abs(forecast)}*20"}, ignore_index=True)
        else:
            beer_second_table = beer_second_table._append({beer_second_table.columns[0]: nomenclature, beer_second_table.columns[1]: f"{abs(forecast)}*30"}, ignore_index=True)

    #Вторая таблица для закусок к пиву
    for index, row in snacks_results.iterrows():
        nomenclature = row['Номенклатура']
        forecast = row['Заказ']

        if forecast == 0:
            snacks_second_table = snacks_second_table._append({snacks_second_table.columns[0]: nomenclature, snacks_second_table.columns[1]: f"{int(forecast*1000)} шт."}, ignore_index=True)
        else:
            snacks_second_table = snacks_second_table._append({snacks_second_table.columns[0]: nomenclature, snacks_second_table.columns[1]: f"{int(custom_ceil(forecast))} кг."}, ignore_index=True)

    #Вторая таблица для прочего
    for index, row in other_results.iterrows():
        nomenclature = row['Номенклатура']
        forecast = row['Заказ']

        other_second_table = other_second_table._append({other_second_table.columns[0]: nomenclature, other_second_table.columns[1]: f"{int(forecast*1000)} шт."}, ignore_index=True)

    save_file()

    # Запись таблиц в один Excel файл
    with pd.ExcelWriter(report_file_path) as writer:
        beer_results.to_excel(writer, sheet_name="Пиво", index=False)
        beer_second_table.to_excel(writer, sheet_name="Пиво", startrow=len(beer_results) + 3, index=False)
        snacks_results.to_excel(writer, sheet_name="Закуски к пиву", index=False)
        snacks_second_table.to_excel(writer, sheet_name="Закуски к пиву", startrow=len(snacks_results) + 3, index=False)
        other_results.to_excel(writer, sheet_name="Прочее", index=False)
        other_second_table.to_excel(writer, sheet_name="Прочее", startrow=len(other_results) + 3, index=False)
        message_result = tk.Label(window, text=f"Файл {report_file_path.split('/')[-1]} загружен", foreground='blue')
        message_result.place(x=185, y=175)

    open_file (report_file_path)


window = tk.Tk()
window.title("Приложение")
window.geometry("500x150")
message_label_sales = tk.Label(window, text="Файл не выбран")
message_label_sales.place(x=65, y=40)
message_label_residuals = tk.Label(window, text="Файл не выбран")
message_label_residuals.place(x=340, y=40)
message_file_not_upload = tk.Label(window, foreground='red')
message_file_not_upload.place(x=185, y=60)

btn_sales = tk.Button(window, text="Загрузить файл с продажами", command=load_sales_file)
btn_residuals = tk.Button(window, text="Загрузить файл с остатками", command=load_residuals_file)
btn_sales.place(x=25, y=10)
btn_residuals.place(x=300, y=10)
btn_generate = tk.Button(window, text="Сгенерировать отчет", command=generate_report)
btn_generate.place(x=185, y=100)

window.mainloop()
