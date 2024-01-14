import tkinter as tk
from collections import defaultdict
from tkinter import messagebox
from datetime import date, datetime

import os
import re

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Приложение для подсчета бандеролей")

        self.total_weight = 0
        self.total_cost = 0.0
        self.total_parcels = 0

        self.weights = []

        # Получить путь к рабочему столу пользователя
        self.desktop_path = os.path.join(os.path.expanduser('~'), 'Desktop')
        
        self.create_main_window()

    def create_main_window(self):
        self.main_window = tk.Toplevel()
        self.main_window.title("Главное окно")
        self.main_window.geometry("600x400")
        
        self.calculate_button = tk.Button(
            self.main_window, text="Подсчет за день", command=self.open_weight_window
        )
        self.calculate_button.pack(pady=10)

        self.letters_button = tk.Button(
            self.main_window, text="Подсчет писем", command=self.calculate_letters
        )
        self.letters_button.pack(pady=10)
        
        self.monthly_button = tk.Button(
            self.main_window, text="Подсчет за месяц", command=self.calculate_total_for_month
        )
        self.monthly_button.pack(pady=10)
        self.monthly_entry_label = tk.Label(self.main_window, text="Введите месяц и год для расчета (мм.гггг):")
        self.monthly_entry_label.pack(pady=10)
        
        self.monthly_entry = tk.Entry(self.main_window)
        self.monthly_entry.pack(pady=10)
        
    def open_weight_window(self):
        self.weight_window = tk.Toplevel(self)
        self.weight_window.title("Подсчет веса бандеролей")
        
        
        # Опцией 'takefocus' задаем перехват фокуса.
        self.weight_window.geometry("400x300")
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
        
        self.delete_weight_button = tk.Button(
            self.weight_window, text="Удалить последний введенный вес", command=self.delete_last_weight
        )
        self.delete_weight_button.pack(pady=10)
        
        self.finish_button = tk.Button(
            self.weight_window, text="Закончить подсчет", command=self.finish_weight_calculation
        )
        self.finish_button.pack(pady=10)
        
        self.weight_entry.bind("<Return>", self.add_weight)

    def add_weight(self, event=None):
        try:
            weight = float(self.weight_entry.get())
            if weight < 120 or weight > 2000:
                raise ValueError("Введите валидный вес (от 120 до 2000).")
            self.weights.append(weight)  # Сначала добавляем вес, только если он валиден
            self.total_weight += weight
            self.total_parcels += 1
            self.total_cost += self.calculate_cost(weight)
            messagebox.showinfo("Успех", "Вес был успешно добавлен.")
            self.weight_entry.delete(0, tk.END)
        except ValueError as e:
            messagebox.showerror("Ошибка", str(e))
        finally:
            self.weight_entry.focus_set()

    
    def delete_last_weight(self):
        if not self.weights:
            messagebox.showinfo("Информация", "Список весов пуст.")
        else:
            last_weight = self.weights.pop()
            self.total_weight -= last_weight
            self.total_parcels -= 1
            self.total_cost -= self.calculate_cost(last_weight)
        
        # Сообщение и фокусировка после изменения
        messagebox.showinfo("Успех", "Последний введенный вес был удален.")
        self.weight_entry.focus_set()

    def finish_weight_calculation(self):
        self.weight_window.destroy()
        self.open_date_window()
        
    def open_date_window(self):
        self.date_window = tk.Toplevel(self)
        self.date_window.title("Ввод даты")
        
        self.date_window.attributes('-topmost', 'true')
        
        self.date_label = tk.Label(self.date_window, text="Введите дату в формате дд.мм.гггг:")
        self.date_label.pack(pady=10)
        
        self.date_entry = tk.Entry(self.date_window)
        self.date_entry.pack(pady=10)
        self.date_entry.focus_set()
        
        self.save_button = tk.Button(
            self.date_window, text="Сохранить результат", command=self.save_results
        )
        self.save_button.pack(pady=10)
        
        self.date_entry.bind("<Return>", self.save_results)

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
        

    def calculate_cost(self, weight):
        base_cost = 89.5
        additional_cost = max(0, (weight - 120) // 20 * 3.5)
        return base_cost + additional_cost
    
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

        self.delete_last_button = tk.Button(self.letters_window, text="Удалить последнее", command=self.remove_last)
        self.delete_last_button.pack(pady=(5,10))

        # Кнопка ОК для запуска подсчета и сохранения результата
        self.ok_button = tk.Button(self.letters_window, text="ОК", command=self.ok_button_pressed)
        self.ok_button.pack(pady=(5,10))

        # Ввод даты
        self.date_label = tk.Label(self.letters_window, text="Введите дату (дд.мм.гггг):")
        self.date_label.pack(pady=(10,5))
        self.date_entry = tk.Entry(self.letters_window)
        self.date_entry.pack(pady=5)
        self.date_entry.bind("<Return>", self.calculate_and_save_result)  # Привязка к кнопке Enter
        
        # Список введенных значений
        self.listbox = tk.Listbox(self.letters_window)
        self.listbox.pack(pady=(10,5))

    def ok_button_pressed(self):
        self.calculate_and_save_result(None)

    def remove_last(self):
        if self.numbers_entered:
            self.numbers_entered.pop()  # Удалить последнее значение
            self.listbox.delete(tk.END)

    
    def add_to_list(self, event):
        # Попытка преобразовать введенные данные в число и добавление в список
        try:
            num_letters = int(self.quantity_entry.get())
            self.numbers_entered.append(num_letters)  # Добавление числа в список
            self.listbox.insert(tk.END, num_letters)  # Вывод числа в интерфейсе
            self.quantity_entry.delete(0, tk.END)  # Очистка поля ввода
        except ValueError:
            tk.messagebox.showwarning("Ошибка", "Введите корректное число!")

    def calculate_and_save_result(self, event):
        # Ввод даты и подсчет итога
        date = self.date_entry.get()
        if not date or not self.numbers_entered:
            tk.messagebox.showwarning("Ошибка", "Введите все данные корректно!")
            return

        # Подсчет итога
        total = sum(self.numbers_entered) * 67  # 67 рублей за письмо
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

    def calculate_total_for_month(self, event=None):
        try:
            selected_month = self.monthly_entry.get()
            # Проверка правильности формата ввода месяца и года
            if not re.match(r"^\d{2}\.\d{4}$", selected_month):
                raise ValueError("Введите месяц и год в правильном формате (мм.гггг).")

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



if __name__ == "__main__":
    app = App()
    app.mainloop()