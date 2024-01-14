def calculate_price(weight):
    base_price = 89.5  # Цена за 120 грамм
    extra_grams = weight - 120  # Количество грамм сверх базового веса
    extra_price = extra_grams // 20 * 3.5  # Стоимость за каждые 20 грамм сверх базового веса
    total_price = base_price + extra_price
    return total_price

daily_total_price = 0
daily_total_parcel = 0
monthly_total_price = 0
monthly_total_parcel = 0

print('Привет, Начальник! Приветствую тебя в приложении подсчета бандероелей! Сейчас ниже появится поле, в которое просто вбивай вес бандероли, после чего жми ENTER. Повторяй данное действие до тех пор, пока не закончатся все бандерольки, а что делать дальше, я подскажу.' '\n' + "Если ввели неверный вес, можно его удалить, введя - 2 и нажав ENTER." + '\n' + "Если бандероли кончились, вводи 0 и жми ENTER ")

weights = [] # Список для хранения введенных весов

while True:
    weight = int(input("Введи вес бандероли (в граммах): " + '\n'))

    if weight == 0:
        print(f"Сегодня было внесено {daily_total_parcel} бандеролей на общую сумму в {daily_total_price} рублей")
        monthly_total_price += daily_total_price
        monthly_total_parcel += daily_total_parcel
        answer = input("Начальник, надо ли добавить что-то? если (да- вводи 1/ нет- вводи 0): ")
        if answer.lower() != "1":
            break
        daily_total_price = 0
        daily_total_parcel = 0
        weights = [] # Очищаем список весов
        continue
    if weight == 2: # Если ввели 2, то удаляем последний введенный вес
        if weights: # Если список весов не пустой
            last_weight = weights.pop() # Удаляем последний элемент и сохраняем его
            last_price = calculate_price(last_weight) # Вычисляем его стоимость
            daily_total_price -= last_price # Вычитаем его из общей суммы
            daily_total_parcel -= 1 # Уменьшаем количество бандеролей на 1
            print(f"Удален вес {last_weight} грамм, стоимостью {last_price} рублей")
        else: # Если список весов пустой
            print("Нет весов для удаления")
        continue
    if weight < 120 or weight > 2000:
        print("Ну уж нет, внимательней, вес должен быть от 120 до 2000 грамм")
        continue
    price = calculate_price(weight)
    daily_total_price += price
    daily_total_parcel += 1
    weights.append(weight) # Добавляем вес в список

    

print(f"Хорошо, за сегодня было внесено {monthly_total_parcel} бандеролей, на общую сумму в {monthly_total_price} рублей.")

date = input("Так, сказал бы что отлично, но кому хочеться возиться с этим го..щем, не правда ли?) А теперь нужно ввести дату формирования списка и снова нажать ENTER: ")
print("Прекрасно, вы ввели дату: ", date)

import os

# Получаем путь к рабочему столу
desktop_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')

# Сохраняем файл на рабочий стол
file_name = "итоги по дням.txt"
file_content = "This is an example file content"
file_path = os.path.join(desktop_path, file_name)

with open(file_path, "a") as file:
    file.write('\n' + str(monthly_total_parcel)+ " " + str(monthly_total_price))

print(f"File {file_name} has been saved to {desktop_path}")

import os
from datetime import datetime

file_name = "итоги по дням.txt"
desktop_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')  # путь до рабочего стола
file_path = None

for root, dirs, files in os.walk(desktop_path):
    if file_name in files:
        file_path = os.path.join(root, file_name)
        break

if file_path:
    with open(file_path, 'a') as file:
        
        file.write(f"\nДата формирования результата: " + date + '\n' + "-----------------------------------------------------------" + "\n")
