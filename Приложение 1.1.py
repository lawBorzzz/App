import tkinter as tk

from tkinter import simpledialog, messagebox
import os

class Application(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.master.geometry("300x150")
        self.pack()
        self.create_widgets()
        self.data = []

    def create_widgets(self):
        self.button1 = tk.Button(self, text="Ввести данные для подсчета", command=self.enter_data)
        self.button1.pack(side="top")


        self.button2 = tk.Button(self, text="Подсчет итога за месяц", command=self.calculate_month_total)
        self.button2.pack(side="top")

    def enter_data(self):
        self.data = []
        self.enter_weight()

    def enter_weight(self):
        weight_prompt = "Введите вес бандероли (от 120 до 2000 грамм):\n\n"\
                        "Если ввели неверный вес, можно его удалить, введя - 1 и нажав ENTER. \n"\
                        "Если бандероли кончились, вводи 0 и жми ENTER."
        weight = simpledialog.askinteger("Ввод веса", weight_prompt, parent=self.master)
        self.enter_weight()
        
        if weight == 0:
            self.calculate_total()
        elif weight == 1:
            if self.data:
                self.data.pop()
                self.enter_weight()
            else:
                messagebox.showwarning("Внимание", "Список весов пуст")
                self.enter_weight()
        elif weight:
            if 120 <= weight <= 2000:
                self.data.append(weight)
                self.enter_weight()
            else:
                messagebox.showerror("Ошибка", "Вес бандероли должен быть от 120 до 2000 грамм")
                self.enter_weight()
        else:
            messagebox.showerror("Ошибка", "Введите корректные данные")
            

    def calculate_total(self):
        total_price = 0
        for weight in self.data:
            total_price += 89.5 + ((weight - 120) // 20) * 3.5

        date = simpledialog.askstring("    Ввод даты    ", "    Введите дату в формате - день.месяц.год:    ", parent=self.master)
        if date:
            file_path = os.path.join(os.path.expanduser("~"), "Desktop", "Итоги по дням.txt")
            with open(file_path, "a") as file:
                file.write(f"За: {date}, количество бандеролей: {len(self.data)}, общая цена: {total_price} рублей\n")
            messagebox.showinfo("Сохранение", "Данные успешно сохранены")

    def calculate_month_total(self):
        # Сначала спрашиваем только месяц и год
        month_year = simpledialog.askstring("    Ввод даты    ", "    Введите месяц и год в формате - месяц.год:    ", parent=self.master)
        if month_year:
            try:
                # Проверяем, что пользователь ввел данные в правильном формате и разделяем их.
                month, year = month_year.split('.')
                if len(month) != 2 or len(year) != 4:
                    raise ValueError("Формат даты должен быть месяц.год (например, 02.2024)")
                

                month_total_banderol = 0
                month_total_price = 0
                data_file_path = os.path.join(os.path.expanduser("~"), "Desktop", "Итоги по дням.txt")

                if not os.path.exists(data_file_path):
                    messagebox.showerror("Ошибка", "Файл с данными не найден")
                    return  # Прекращаем выполнение функции, если файл отсутствует

                with open(data_file_path, "r") as file:
                    for line in file:
                        # Получаем данные даты в правильном формате из строки
                        date_str = line[line.index("За: ") + 6:line.index(',')]
                        day, line_month, line_year = date_str.split('.')
                        # Проверяем, что строка соответствует заданному месяцу и году
                        if line_month == month and line_year == year:
                            info = line.split(", ")
                            for data in info:
                                if "количество бандеролей" in data:
                                    month_total_banderol += int(data.split(":")[1])
                                elif "общая цена" in data:
                                    month_total_price += float(data.split(": ")[1].split()[0])

                # Запись итогов за месяц в отдельный файл
                file_path = os.path.join(os.path.expanduser("~"), "Desktop", "Итоги за месяц.txt")
                with open(file_path, "a") as file:
                    file.write(
                        f"Итог за {month_year}: общее количество бандеролей - {month_total_banderol}, общая цена - {month_total_price} рублей\n")
                messagebox.showinfo("Сохранение", "Итог за месяц успешно сохранен")
            except ValueError as e:
                messagebox.showerror("Ошибка", str(e))

root = tk.Tk()
app = Application(master=root)
app.master.title("Расчет бандеролей")
app.mainloop()
#version 1.1
