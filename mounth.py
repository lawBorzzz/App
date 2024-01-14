total_price = 0
total_parcel = 0

import os

file_name = "итоги по дням.txt"
desktop_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')  # путь до рабочего стола

for root, dirs, files in os.walk(desktop_path):
    if file_name in files:
        file_path = os.path.join(root, file_name)
        print("Найден файл по пути:", file_path)


with open(file_path, 'r') as file:
    for line in file:
        # Разделяем строку на количество и цену, предполагая, что они разделены пробелом
        values = line.split()
        
        if len(values) == 2:  # Предполагаем, что строка содержит количество и цену
            count = int(values[0])
            price = float(values[1])
            
            # Считаем сумму
            total_price += price
            total_parcel += count

# Выводим итог
print("Общая сумма:", total_price)
print("Общее количество бандеролей:", total_parcel)

date = input("Так, a теперь нужно ввести дату формирования списка и снова нажать ENTER: ")
print("Прекрасно, вы ввели дату: ", date)

import os

# Получаем путь к рабочему столу
desktop_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')

# Сохраняем результат в файл на рабочий стол
file_name = "итог за месяц.txt"
file_path = os.path.join(desktop_path, file_name)

with open(file_path, "a") as file:
    file.write("Итог за месяц: Всего бандеролей: " + str(total_parcel) + " Общая цена: " + str(total_price) + " рублей\n")

print(f"Результат был сохранен в файл {file_name} на рабочем столе" + "\n")

import os
from datetime import datetime

file_name = "итог за месяц.txt"
desktop_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')  # путь до рабочего стола
file_path = None

for root, dirs, files in os.walk(desktop_path):
    if file_name in files:
        file_path = os.path.join(root, file_name)
        break

if file_path:
    with open(file_path, 'a') as file:
        
        file.write(f"\nДата формирования результата: " + date + '\n' + "-----------------------------------------------------------" + "\n")
