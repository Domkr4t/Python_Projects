import json
import csv
import time

start_time = time.time()

# Load the JSON data
with open('data.json', 'r', encoding='utf-8') as json_file:
    data = json.load(json_file)

# Prepare the CSV file
with open('data.csv', 'w', encoding='utf-8', newline='') as csv_file:
    csv_writer = csv.writer(csv_file)

    # Write the headers
    headers = data[0].keys()
    csv_writer.writerow(headers)

    # Write the data
    for item in data:
        csv_writer.writerow(item.values())


end_time = time.time() # Запись времени окончания выполнения
execution_time = end_time - start_time # Вычисление общего времени выполнения

print(f"Время выполнения программы: {execution_time} секунд")
print("Все супер!")
