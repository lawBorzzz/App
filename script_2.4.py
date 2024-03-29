import tkinter as tk
from collections import defaultdict
from tkinter import messagebox
from datetime import date, datetime

import os
import re
import math

class App(tk.Tk):

    BASE_COST = 89.5  # базовая стоимость бандероли
    STEP_COST = 3.5   # стоимость за шаг в 20 грамм
    LETTER_COST = 67  # стоимость письма

# Это округление введенных бандеролей до целого четного числа равному 20
    def round_weight(self, weight):
        return math.ceil(weight / 20.0) * 20


    def __init__(self):
        super().__init__()
        self.title("Приложение для подсчета бандеролей")

        self.total_weight = 0
        self.total_cost = 0.0
        self.total_parcels = 0

        self.weights = []

        # Получить путь к рабочему столу пользователя
        self.desktop_path = os.path.join(os.path.expanduser('~'), 'Desktop')
        
        self.create_widgets()

# Это главное окно с последующими кнопками
    def create_widgets(self):
        self.geometry("600x400")  # Прямо здесь задаем размер корневого окна
        
        # Кнопка подсчета за день
        self.calculate_button = tk.Button(
            self, text="Подсчет за день", command=self.open_weight_window
        )
        self.calculate_button.pack(pady=10)
        
        # Кнопка подсчета писем
        self.letters_button = tk.Button(
            self, text="Подсчет писем", command=self.calculate_letters
        )
        self.letters_button.pack(pady=10)
        
        # Кнопка подсчета за месяц
        self.monthly_button = tk.Button(
            self, text="Подсчет за месяц", command=self.ask_month_input  # Изменено здесь
        )
        self.monthly_button.pack(pady=10)

        # Кнопка настроек
        self.settings_button = tk.Button(
            self, text="Настройки", command=self.open_settings_window
        )
        self.settings_button.pack(pady=10)

# Функция для создания диалогового окна и ввода даты
    def ask_month_input(self):

        self.month_window = tk.Toplevel(self)
        self.month_window.title("Введите месяц и год")
        self.month_window.geometry("400x150")
        self.month_window.attributes('-topmost', 'true')
        
        self.instruction_label = tk.Label(self.month_window, text="Введите дату для расчета за период в формате мм.гггг:")
        self.instruction_label.pack(pady=(10, 0))

        self.month_entry = tk.Entry(self.month_window)
        self.month_entry.pack(pady=10)
        self.month_entry.focus_set()

        self.ok_button = tk.Button(self.month_window, text="OK", command=self.initiate_monthly_calculation)  # Создаем кнопку OK
        self.ok_button.pack()

        self.month_window.bind("<Return>", self.initiate_monthly_calculation)  # Изменено здесь - привязываем к окну ввода


    def initiate_monthly_calculation(self, event=None):
        # Это новая функция, которая будет вызывать calculate_total_for_month
        self.calculate_total_for_month()
        self.month_window.destroy()

# Открывается кнопка Настроек
    def open_settings_window(self):
        self.settings_window = tk.Toplevel(self)
        self.settings_window.title("Настройки")
        
        self.settings_window.attributes('-topmost', 'true')
        
        self.base_cost_label = tk.Label(self.settings_window, text="Стоимость бандероли:")
        self.base_cost_label.pack(pady=10)
        
        self.base_cost_entry = tk.Entry(self.settings_window)
        self.base_cost_entry.pack(pady=10)
        self.base_cost_entry.insert(0, str(self.BASE_COST))
        
        self.step_cost_label = tk.Label(self.settings_window, text="Стоимость за шаг в 20 грамм:")
        self.step_cost_label.pack(pady=10)
        
        self.step_cost_entry = tk.Entry(self.settings_window)
        self.step_cost_entry.pack(pady=10)
        self.step_cost_entry.insert(0, str(self.STEP_COST))
        
        self.letter_cost_label = tk.Label(self.settings_window, text="Стоимость письма:")
        self.letter_cost_label.pack(pady=10)
        
        self.letter_cost_entry = tk.Entry(self.settings_window)
        self.letter_cost_entry.pack(pady=10)
        self.letter_cost_entry.insert(0, str(self.LETTER_COST))
        
        self.save_settings_button = tk.Button(
            self.settings_window, text="Сохранить", command=self.save_settings
        )
        self.save_settings_button.pack(pady=10)

    def save_settings(self):
        try:
            self.BASE_COST = float(self.base_cost_entry.get())
            self.STEP_COST = float(self.step_cost_entry.get())
            self.LETTER_COST = float(self.letter_cost_entry.get())
            self.settings_window.destroy()
            messagebox.showinfo("Успех", "Настройки сохранены.")
        except ValueError:
            messagebox.showerror("Ошибка", "Введите корректные числовые значения.")

# ОТкрывается кнопка ввода веса бандеролей в день        
    def open_weight_window(self):
        self.weight_window = tk.Toplevel(self)
        self.weight_window.title("Подсчет веса бандеролей")
        
        self.weight_window.geometry("400x500")
        self.weight_window.attributes('-topmost', 'true')
        
        
        self.weight_label = tk.Label(self.weight_window, text="Введите вес бандероли (в граммах):")
        self.weight_label.pack(pady=10)
        
        self.weight_entry = tk.Entry(self.weight_window)
        self.weight_entry.pack(pady=10)
        self.weight_entry.focus_set()
        
        self.add_weight_button = tk.Button(
            self.weight_window, text="Добавить вес", command=self.add_weight
        )
        self.add_weight_button.pack(pady=10)
        self.weight_entry.focus_set()

        # Создание Listbox для отображения введенных весов
        self.weights_listbox_label = tk.Label(self.weight_window, text="Введенные веса:")
        self.weights_listbox_label.pack()
        self.weights_listbox = tk.Listbox(self.weight_window)
        self.weights_listbox.pack()
        
        self.delete_weight_button = tk.Button(
            self.weight_window, text="Удалить последний введенный вес", command=self.delete_last_weight
        )
        self.delete_weight_button.pack(pady=10)
        
        self.finish_button = tk.Button(
            self.weight_window, text="Закончить подсчет", command=self.finish_weight_calculation
        )
        self.finish_button.pack(pady=10)
        
        self.weight_entry.bind("<Return>", self.add_weight)


# Добавление бандеролей в общий список до подсчета
    def add_weight(self, event=None):
        try:
            weight = float(self.weight_entry.get())
            # Округляем вес
            rounded_weight = self.round_weight(weight)
            if rounded_weight < 120 or rounded_weight > 2000:
                raise ValueError("Введите валидный вес (от 120 до 2000).")
            self.weights.append(rounded_weight)  # Добавляем округленный вес
            self.total_weight += rounded_weight  # Используем округленный вес для общего веса
            self.total_parcels += 1
            self.total_cost += self.calculate_cost(rounded_weight)  # Рассчитываем стоимость по округленному весу
            self.weights_listbox.insert(tk.END, f"{rounded_weight} грамм")
            self.weight_entry.delete(0, tk.END)
        except ValueError as e:
            messagebox.showerror("Ошибка", str(e))
        finally:
            self.weight_entry.focus_set()

# Удаление крайнего веса из списка   
    def delete_last_weight(self):
        if not self.weights:
            messagebox.showinfo("Информация", "Список весов пуст.")
        else:
            last_weight = self.weights.pop()
            self.total_weight -= last_weight
            self.total_parcels -= 1
            self.total_cost -= self.calculate_cost(last_weight)
            self.weights_listbox.delete(tk.END)
        
        # Сообщение и фокусировка после изменения
        self.weight_entry.focus_set()

    def finish_weight_calculation(self):
        self.weight_window.destroy()
        self.open_date_window()



# Открывается поле ввода даты перед сохранением списка с бандеролями       
    def open_date_window(self):
        self.date_window = tk.Toplevel(self)
        self.date_window.title("Ввод даты")
        self.date_window.geometry("300x150")
        self.date_window.attributes('-topmost', 'true')
        
        self.date_label = tk.Label(self.date_window, text="Введите дату в формате дд.мм.гггг:")
        self.date_label.pack(pady=10)
        
        self.date_entry = tk.Entry(self.date_window)
        self.date_entry.pack(pady=10)
        self.date_entry.focus_set()
        
        self.save_button = tk.Button(
            self.date_window, text="Сформировать список", command=self.save_results
        )
        self.save_button.pack(pady=10)
        
        self.date_entry.bind("<Return>", self.save_results)

# Сохранение результата списка бандеролей с датой
    def save_results(self, event=None):
        try:
            entry_date = self.date_entry.get()
            current_date = datetime.strptime(entry_date, "%d.%m.%Y").date()
            
            result_string = (f"Итого за {current_date.strftime('%d.%m.%Y')} отправлено {self.total_parcels}"
                             f" бандеролей весом {self.total_weight:.2f} грамм на сумму {self.total_cost:.2f} рублей.\n")
            
            filepath = os.path.join(self.desktop_path, "итого.txt")
            with open(filepath, "a", encoding='utf-8') as file:
                file.write(result_string)
            
            messagebox.showinfo("Успех", "Результаты сохранены.")
            self.date_window.destroy()
            
            self.weights.clear()
            self.total_weight = 0
            self.total_cost = 0.0
            self.total_parcels = 0
        except ValueError:
            messagebox.showerror("Ошибка", "Введите дату в правильном формате (дд.мм.гггг).")
        
# Расчет стоимости бандеролей, идет по настройкам с выставленными значениями.
    def calculate_cost(self, weight):
        additional_cost = max(0, (weight - 120) // 20 * self.STEP_COST)
        return self.BASE_COST + additional_cost
    
# Окно с полем ввода количества писем   
    def calculate_letters(self):
        # Открытие нового окна для ввода количества писем и даты
        self.letters_window = tk.Toplevel()
        self.letters_window.title("Подсчет писем")
        self.letters_window.geometry("300x600")

        self.numbers_entered = []  # Список для хранения введенных значений
        
        # Ввод количества писем
        self.quantity_label = tk.Label(self.letters_window, text="Введите количество писем и нажмите Enter:")
        self.quantity_label.pack(pady=(20,5))
        
        self.quantity_entry = tk.Entry(self.letters_window)
        self.quantity_entry.pack(pady=5)
        self.quantity_entry.bind("<Return>", self.add_to_list)  # Привязка к кнопке Enter
        
        # Список введенных значений
        self.listbox_label = tk.Label(self.letters_window, text="Список введённых значений:")
        self.listbox_label.pack(pady=(10,0))  # Отступ сверху перед надписью
        self.listbox = tk.Listbox(self.letters_window)
        self.listbox.pack(pady=(0,5))

        # Кнопка "удалить последнее" для удаления последнегго введенного результата
        self.delete_last_button = tk.Button(self.letters_window, text="Удалить последнее", command=self.remove_last)
        self.delete_last_button.pack(pady=(5,10))

        # Ввод даты
        self.date_label = tk.Label(self.letters_window, text="Введите дату (дд.мм.гггг):")
        self.date_label.pack(pady=(10,5))
        self.date_entry = tk.Entry(self.letters_window)
        self.date_entry.pack(pady=5)
        self.date_entry.bind("<Return>", self.calculate_and_save_result)  # Привязка к кнопке Enter

        # Кнопка ОК для запуска подсчета и сохранения результата
        self.ok_button = tk.Button(self.letters_window, text="Сформировать список", command=self.ok_button_pressed)
        self.ok_button.pack(pady=(5,10))

# После нажатия кнопки ОК, считаем и сохраняем результат
    def ok_button_pressed(self):
        self.calculate_and_save_result(None)

# Если нажали кнопку удалить последний результат из списка писем.
    def remove_last(self):
        if self.numbers_entered:
            self.numbers_entered.pop()  # Удалить последнее значение
            self.listbox.delete(tk.END)

# Это лист, где отображаются введенные письма (как в памяти так и в окне в виде списка)    
    def add_to_list(self, event):
        # Попытка преобразовать введенные данные в число и добавление в список
        try:
            num_letters = int(self.quantity_entry.get())
            self.numbers_entered.append(num_letters)  # Добавление числа в список
            self.listbox.insert(tk.END, num_letters)  # Вывод числа в интерфейсе
            self.quantity_entry.delete(0, tk.END)  # Очистка поля ввода
        except ValueError:
            tk.messagebox.showwarning("Ошибка", "Введите корректное число!")
# Подсчет и сохранение итога по письмам
    def calculate_and_save_result(self, event):
        # Ввод даты и подсчет итога
        date = self.date_entry.get()
        if not date or not self.numbers_entered:
            tk.messagebox.showwarning("Ошибка", "Введите все данные корректно!")
            return

        # Подсчет итога
        total = sum(self.numbers_entered) * self.LETTER_COST
        self.save_to_file(total, sum(self.numbers_entered), date)

        # Отображение результата
        tk.messagebox.showinfo("Результаты подсчета", f"Итого количество писем: {sum(self.numbers_entered)}\nНа сумму: {total} руб.\n")
    
        # Закрытие окна ввода
        self.letters_window.destroy()

    def save_to_file(self, total_result, total_letters, date):
        desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
        filename = os.path.join(desktop, f"итоги писем.txt")

        # Используем режим 'a' для добавления данных в конец файла
        with open(filename, 'a', encoding='utf-8') as file:
            file.write(f"\nДата: {date}\nКоличество писем: {total_letters}\nИтого: {total_result} руб.\n")

        # Оповещение пользователя, что сохранение прошло успешно
        tk.messagebox.showinfo("Сохранение результата", "Результат успешно сохранен в файл:\n"+filename)


# Это сохранение итогов бандеролей по конкретному периоду (за определенный месяц)
    def calculate_total_for_month(self, event=None):
        try:
            selected_month = self.month_entry.get()
            # Проверка правильности формата ввода месяца и года
            if not re.match(r"^\d{2}\.\d{4}$", selected_month):
                raise ValueError("Введите месяц и год в правильном формате (мм.гггг).")
            # Если ввод прошел успешно, закрыть окно ввода
            
            # Остальная логика функции по подсчету итогов за месяц...

            month, year = map(int, selected_month.split("."))
        
            total_weight = 0
            total_cost = 0.0
            total_parcels = 0
    
            filepath = os.path.join(self.desktop_path, "итого.txt")
            # Проверка на существование файла
            if not os.path.exists(filepath):
                raise FileNotFoundError(f"Файл {filepath} не найден.")

            with open(filepath, 'r', encoding='utf-8') as file:
                for line in file:
                    # Проверка соответствия месяца и года в строке с выбранными
                    date_match = re.search(r'(\d{2})\.(\d{2})\.(\d{4})', line)
                    if date_match:
                        line_day, line_month, line_year = map(int, date_match.groups())
                        if line_month == month and line_year == year:
                            # Разбор данных о количестве и весе бандеролей
                            total_parcels += int(re.search(r'отправлено (\d+) бандеролей', line).group(1))
                            total_weight += float(re.search(r'весом ([\d.]+) грамм', line).group(1))
                            total_cost += float(re.search(r'на сумму ([\d.]+) рублей', line).group(1))
        
            # Строка результата расчета
            result_string = (f"Итого за {selected_month}: отправлено {total_parcels} бандеролей "
                             f"весом {total_weight} грамм на сумму {total_cost} рублей.\n")

            # Показываем результат в message box
            messagebox.showinfo("Итоги за месяц", result_string)
    
            # Сохранение в файл (при необходимости)
            output_file_path = os.path.join(self.desktop_path, f"итог за {selected_month}.txt")
            with open(output_file_path, 'w', encoding='utf-8') as output_file:
                output_file.write(result_string)
            
        except ValueError as ve:
            messagebox.showerror("Ошибка", str(ve))
        except FileNotFoundError as fe:
            messagebox.showerror("Ошибка", str(fe))

# конец(((

if __name__ == "__main__":
    app = App()
    app.mainloop()