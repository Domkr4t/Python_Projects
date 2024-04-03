import json
import time

start_time = time.time()

# Load the JSON data
with open('output.json', 'r', encoding='utf-8') as json_file:
    data = json.load(json_file)

# Modify the data
for item in data:
    if item.get("Категория товара") == "Рыба":
        item["Категория товара"] = "Закуски к пиву"

# Save the modified data
with open('data.json', 'w', encoding='utf-8') as json_file:
    json.dump(data, json_file, ensure_ascii=False, indent=4)


end_time = time.time() # Запись времени окончания выполнения
execution_time = end_time - start_time # Вычисление общего времени выполнения

print(f"Время выполнения программы: {execution_time} секунд")
print("Все супер!")
