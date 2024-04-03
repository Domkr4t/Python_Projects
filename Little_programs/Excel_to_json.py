import pandas as pd
import json
import time

# Переменная для названий файлов
i = 1

start_time = time.time() # Запись времени начала выполнения программы

while (i < 5):
    # Загрузка данных из Excel файла
    file_path = f'C:\\Users\\User\\Desktop\\import{i}.xlsx'
    try:
        #Пропускаем первые 5 строк, и берем только 1, 4, 5, 7... столбцы
        df_new = pd.read_excel(file_path, skiprows=4, usecols=[0, 3, 4, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17])
    except FileNotFoundError:
        print(f"Файл {file_path} не найден.")
        i += 1
        continue

    # Преобразование столбца с датами в строку
    df_new['Дата'] = df_new['Дата'].astype(str)

    # Переименование столбцов
    # df_new = df_new.rename(columns={
    #     'Магазин': 'shop',
    #     'Дата': 'dateOfBuy',
    #     'Номер чека': 'checkNumber',
    #     'Дисконтная карта': 'discountCard',
    #     'Владелец карты': 'cardholder',
    #     'Номер телефона': 'phoneNumber',
    #     'Категория товара': 'productCategory',
    #     'Группа товара': 'productGroup',
    #     'Номенклатура': 'nomenclature',
    #     'Сумма продаж': 'salesAmount',
    #     'Количество': 'quantity',
    #     'Остаток на складе': 'stockBalance',
    #     'Сумма скидки': 'discountAmount',
    #     'Себестоимость': 'costPrice',
    #     'Валовая прибыль': 'grossProfit',
    # })

    # Чтение существующего JSON файла
    try:
        with open('C:\\Users\\User\\Desktop\\output.json', 'r', encoding='utf-8') as f:
            data_existing = json.load(f)
    except FileNotFoundError:
        print("Файл output.json не найден. Создание нового файла.")
        data_existing = []

    # Преобразование списка словарей в DataFrame
    df_existing = pd.DataFrame(data_existing)

    # Объединение двух DataFrame в один
    df = pd.concat([df_existing, df_new], ignore_index=True)

    # Преобразование объединенного DataFrame в JSON
    data = df.to_dict(orient='records')

    # Запись данных обратно в JSON файл
    with open('C:\\Users\\User\\Desktop\\output.json', 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=4)

    print(f"Файл import{i}.xlsx записан.")

    i += 1


end_time = time.time() # Запись времени окончания выполнения
execution_time = end_time - start_time # Вычисление общего времени выполнения

print(f"Время выполнения программы: {execution_time} секунд")
print("Данные успешно записаны в файл output.json")
