import pandas as pd
import os
import math
import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
import platform


sales_file_path = None
residuals_file_path = None
report_file_path = None
category_var = None

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


def handle_snacks(nomenclature, stock, total_quantity):
    forecasted_balance = stock - total_quantity
    if (forecasted_balance < 0.600 and not forecasted_balance.is_integer()):
        forecast = 1 - abs(forecasted_balance)
    if (forecasted_balance < 0 and not forecasted_balance.is_integer()):
        forecast = abs(forecasted_balance)
    elif (forecasted_balance >= 0.600 and not forecasted_balance.is_integer()):
        forecast = 0
    elif (forecasted_balance.is_integer()):
        forecast = 0

    return {'Номенклатура': nomenclature, 'Остаток': stock, 'Прогноз': total_quantity, 'Прогнозируемый остаток': forecasted_balance, 'Заказ': abs(forecast)}


def handle_other(nomenclature, stock, total_quantity):
    forecasted_balance = stock - total_quantity
    if (forecasted_balance < 600 and not forecasted_balance.is_integer()):
        forecast = 1 - forecasted_balance
    elif (forecasted_balance >= 600 and not forecasted_balance.is_integer()):
        forecast = 0
    elif (forecasted_balance.is_integer()):
        forecast = 0

    return {'Номенклатура': nomenclature, 'Остаток': stock, 'Прогноз': total_quantity, 'Прогнозируемый остаток': forecasted_balance, 'Заказ': abs(forecast)}


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

def open_file(file_path):
    # Открывает файл с помощью системной команды.
    if platform.system() == "Windows":
        os.startfile(file_path)
    elif platform.system() == "Darwin":
        os.system(f"open {file_path}")
    else:
        os.system(f"xdg-open {file_path}")

def save_file(category):
    global report_file_path
    # Открываем диалог выбора файла для сохранения
    report_file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", title=f'Сохранить отчет "{category}"', filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")])


def generate_report():
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

    print(sales.columns)


    # Чтение данных из второго файла (с остатками)
    residuals = pd.read_excel(f'{residuals_file_path}', skiprows=2, usecols=[0, 2]) # пропуск строки с назва
    residuals_column_names = ['Номенклатура', 'Остаток']
    residuals.columns = residuals_column_names

    print(residuals.columns)

    selected_category = category_var.get()

    if selected_category == "Пиво":
        results = pd.DataFrame(columns=['Номенклатура', 'Остаток', 'Прогноз', 'Заказ кег', 'Остаток литров'])
        second_table = pd.DataFrame(columns=['Номенклатура', 'Заказ кег'])
    elif selected_category == "Закуски к пиву":
        results = pd.DataFrame(columns=['Номенклатура', 'Остаток', 'Прогноз', 'Прогнозируемый остаток', 'Заказ'])
        second_table = pd.DataFrame(columns=['Номенклатура', 'Заказ'])
    elif selected_category == "Прочее":
        results = pd.DataFrame(columns=['Номенклатура', 'Остаток', 'Прогноз', 'Прогнозируемый остаток', 'Заказ'])
        second_table = pd.DataFrame(columns=['Номенклатура', 'Заказ'])

    # Обработка данных
    for index, row in residuals.iterrows():

        nomenclature = row['Номенклатура']
        stock = row['Остаток']

        # Находим все записи в первом файле, соответствующие текущей номенклатуре
        matching_rows = sales[sales['Номенклатура'] == nomenclature]
        if selected_category == "Пиво":
            if not matching_rows.empty and matching_rows['Категория товара'].str.contains("Пиво").any():
                total_quantity = matching_rows['Количество товара'].sum()

                result = handle_beer(nomenclature, stock, total_quantity)
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
        elif selected_category == "Закуски к пиву":
            if not matching_rows.empty and matching_rows['Категория товара'].str.contains("Закуски к пиву").any() | matching_rows['Категория товара'].str.contains("Рыба").any():
                total_quantity = matching_rows['Количество товара'].sum()

                result = handle_snacks(nomenclature, stock, total_quantity)
                results = results._append(result, ignore_index=True)
        elif selected_category == "Прочее":
            if not matching_rows.empty and matching_rows['Категория товара'].str.contains("Прочее").any():
                total_quantity = matching_rows['Количество товара'].sum()

                result = handle_other(nomenclature, stock, total_quantity)
                results = results._append(result, ignore_index=True)

    #Вторая таблица
    if selected_category == "Пиво":
        for index, row in results.iterrows():
            nomenclature = row['Номенклатура']
            forecast = row['Заказ кег']

            if nomenclature == "Вайлд Черри":
                second_table = second_table._append({second_table.columns[0]: nomenclature, second_table.columns[1]: f"{abs(forecast)}*20"}, ignore_index=True)
            else:
                second_table = second_table._append({second_table.columns[0]: nomenclature, second_table.columns[1]: f"{abs(forecast)}*30"}, ignore_index=True)
    elif selected_category == "Закуски к пиву":
        for index, row in results.iterrows():
            nomenclature = row['Номенклатура']
            forecast = row['Заказ']
            forecasted_balance = row['Прогнозируемый остаток']

            if forecasted_balance.is_integer():
                second_table = second_table._append({second_table.columns[0]: nomenclature, second_table.columns[1]: f"{int(forecast*1000)} шт."}, ignore_index=True)
            else:
                second_table = second_table._append({second_table.columns[0]: nomenclature, second_table.columns[1]: f"{int(custom_ceil(forecast))} кг."}, ignore_index=True)
    elif selected_category == "Прочее":
        for index, row in results.iterrows():
            nomenclature = row['Номенклатура']
            forecast = row['Заказ']

            second_table = second_table._append({second_table.columns[0]: nomenclature, second_table.columns[1]: f"{int(forecast*1000)} шт."}, ignore_index=True)

    save_file(selected_category)

    # Запись двух таблиц в один Excel файл
    with pd.ExcelWriter(report_file_path) as writer:
        results.to_excel(writer, sheet_name=selected_category, index=False)
        second_table.to_excel(writer, sheet_name=selected_category, startrow=len(results) + 3, index=False)
        message_result.config(text=f"Файл {report_file_path.split('/')[-1]} загружен", foreground='blue')


    open_file (report_file_path)


window = tk.Tk()
window.title("Приложение")
window.geometry("500x250")
message_label_sales = tk.Label(window, text="Файл не выбран")
message_label_sales.place(x=65, y=40)
message_label_residuals = tk.Label(window, text="Файл не выбран")
message_label_residuals.place(x=340, y=40)
message_file_not_upload = tk.Label(window, foreground='red')
message_file_not_upload.place(x=185, y=60)
message_result = tk.Label(window, text="")
message_result.place(x=185, y=175)

# Создание ComboBox для выбора категории
category_var = tk.StringVar(window)
category_var.set("Пиво") # Установка значения по умолчанию
category_combobox = ttk.Combobox(window, textvariable=category_var, values=["Пиво", "Закуски к пиву", "Прочее"], state='readonly')
category_combobox.place(x=185, y=100)

btn_sales = tk.Button(window, text="Загрузить файл с продажами", command=load_sales_file)
btn_residuals = tk.Button(window, text="Загрузить файл с остатками", command=load_residuals_file)
btn_sales.place(x=25, y=10)
btn_residuals.place(x=300, y=10)
btn_generate = tk.Button(window, text="Сгенерировать отчет", command=generate_report)
btn_generate.place(x=185, y=200)


window.mainloop()
