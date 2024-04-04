import pandas as pd
import os
import math

#Функция округления
def custom_ceil(number):
    if number < 0:
        return math.floor(number)
    else:
        return math.ceil(number)


# Функция для обработки пива
def handle_beer(nomenclature, stock, total_quantity):
    forecast = math.floor((stock - total_quantity) / 30)
    if (forecast > 0):
        remaining_liters = (stock - total_quantity) % 30
    else:
        remaining_liters = 0
    return {'Номенклатура': nomenclature, 'Остаток': stock, 'Прогноз': total_quantity, 'Заказ кег': forecast, 'Остаток литров': remaining_liters}

sales_path_name = input("Введи название файла с продажами: ")

# Чтение данных из первого файла (с продажами)
sales = pd.read_excel(f'{sales_path_name}.xlsx', skiprows=4, usecols=[0, 3, 4, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17])

sales_column_names = ['Магазин', 'Дата и время', 'Id чека', 'Диск карт', 'Владелец карты', 'Номер телефона', 'Категория товара', 'Группа товара', 'Номенклатура', 'Сумма продаж', 'Количество товара', 'Остаток на складе', 'Сумма скидки', 'Себестоимость продаж', 'Валовая прибыль']
sales.columns = sales_column_names

print(sales.columns)


residuals_path_name = input("Введи название файла с остатками: ")

# Чтение данных из второго файла (с остатками)
residuals = pd.read_excel(f'{residuals_path_name}.xlsx', skiprows=2, usecols=[0, 2]) # пропуск строки с назва
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
if os.path.exists('Прогноз_пиво.xlsx'):
    # Если файл существует, удаляем его
    os.remove('Прогноз_пиво.xlsx')

results.to_excel('Прогноз_пиво.xlsx', index=False)

print("Все успешно!")
