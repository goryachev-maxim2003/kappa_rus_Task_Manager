# pyinstaller --onefile --windowed TD.py //Для того, чтобы сделать exe файл без консоли

import tkinter as tk
import xlwings as xw
import pandas as pd
import numpy as np
import os

# Открытие файла с ответами
# Поиск файла с ответами в директории

files = os.listdir()
# Ищем файл со строкой 'MenedzherZadach' в названии
answers_file_name = ''
for file_name in files:
    if 'MenedzherZadach' in file_name:
        answers_file_name = file_name
        break
text_fault	= "Сообщить о проблеме/аномалии/ неисправности"

answers = pd.read_excel(answers_file_name, index_col="ID")
# Удаление пустых столбцов

problems = answers[answers["Выберите задачу"] == text_fault].dropna(how="all", axis=1)

problems["Время создания"] = pd.to_datetime(problems["Время создания"])

# Соответствие между участком и журналом
optional_fields_in_machines = np.array(["Опишите аномалию", "Вставьте фотографию"])
# Журналы и необязательные поля участка
plot_fields = pd.DataFrame([
    ('Гофроагрегат', 'Журнал неисправностей Corrugator BHS.xlsx', optional_fields_in_machines),
    ('Конвертация', 'Журнал неисправностей Упаковка Конвертации.xlsx', optional_fields_in_machines),
    ('Зона упаковки', 'Журнал неисправностей Упаковка Конвертации.xlsx', optional_fields_in_machines),
    ('Макулатурный участок', 'Журнал неисправностей Участка Макулатуры.xlsx', optional_fields_in_machines),
    ('Непроизводственное оборудование', 'Журнал неисправностей Прочая переферия.xlsx', optional_fields_in_machines),
    ('Мартин 616', 'Журнал неисправностей Martin 616.xlsx', optional_fields_in_machines),
    ('Мартин 924', 'Журнал неисправностей Martin 924.xlsx', optional_fields_in_machines),
    ('Мартин 1232', 'Журнал неисправностей Martin 1232.xlsx', optional_fields_in_machines),
    ('Бобст', 'Журнал неисправностей Bobst.xlsx', optional_fields_in_machines),
    ('Асахи', 'Журнал неисправностей Goepfert-ASAHI.xlsx', optional_fields_in_machines),
    ('Гепферт', 'Журнал неисправностей Goepfert RDC.xlsx', optional_fields_in_machines),
    ('Танабэ', 'Журнал неисправностей Tanabe JD BoxR 1450.xlsx', optional_fields_in_machines)
], columns=["Станок", "Журнал", "Необязательные поля"])
plot_fields.set_index("Станок", inplace=True)

# Распределяем данные по файлам
def get_max_datatime(ser): #Максимальное время в series (для пустой series возвращает самую старую дату)
    return pd.to_datetime(0) if len(ser) == 0 else ser.max()
# Функция записи определённых ответов в журнал определённого участка
def write_in_plot(plot):
    global  journals_books
    errors.delete('1.0', 'end')
    journal = pd.read_excel(plot_fields.loc[plot, "Журнал"], skiprows=[1])
    # Объединяем дату и время
    journal["datatime"] = journal["Дата"] + journal["Время"].apply(lambda t: pd.to_timedelta(str(t)))
    # Находим проблемы определённого участка
    plot_problems = problems[(problems["Выберите участок"] == plot) | (problems["Выберите станок"] == plot)]
    # Выбираем новые записи (те которых не было в файле)
    plot_problems = plot_problems[plot_problems["Время создания"] > get_max_datatime(journal["datatime"])]
    #Заполняем прочерками необязательные поля
    for col in plot_fields.loc[plot, "Необязательные поля"]:
        plot_problems.loc[plot_problems[col].isna(), col] = "-"
    #Перемещаем ФИО, создаём дату и время
    first_column = plot_problems.pop('ФИО')
    plot_problems.insert(1, "ФИО", first_column)
    plot_problems.insert(0, "Дата", plot_problems["Время создания"].dt.date)
    plot_problems["Время создания"] = plot_problems["Время создания"].dt.time.astype(str)
    #Удаляем лишнюю колонку и лишние пустые значения
    plot_problems = plot_problems.drop(columns = ["Выберите задачу", 'Выберите участок', "Выберите станок"], axis = 1)
    plot_problems = plot_problems.apply(lambda row : row.dropna().reset_index(drop = True), axis = 1)
    #Записываем файл
    if (len(plot_problems) > 0):
        book = xw.Book(plot_fields.loc[plot, "Журнал"])
        sht = book.sheets['Sheet1']
        first_empty_row =  3 if (sht.range('A3').value is None) else sht.range('A3').end('down').row + 1
        sht.range(f'A{first_empty_row}').expand(mode='table').value = plot_problems.values
        #Сохраняем файл
        book.save()
    else:
        errors.insert(1.0, "Новых записей не появилось")

def load():
    plots = problems["Выберите участок"].dropna().unique()
    plots = np.concatenate((plots[plots != 'Конвертация'], problems["Выберите станок"].dropna().unique()))
    for plot in plots:
        write_in_plot(plot)
def close():
    errors.delete('1.0', 'end')
    for book in journals_books:
        book.close()
journals_books = []
root = tk.Tk()
root.title("Заполнение журналов на основе Яндекс форм")
root.geometry("500x500")
bt_load = tk.Button(root, text="Загрузить новые данные из yandex forms", width=50, height=1, command=lambda : load())
bt_load.place(x=50, y=25)
bt_close = tk.Button(root, text="Закрыть все журналы", width=50, height=1, command=lambda : close())
bt_close.place(x=50, y=55)
errors = tk.Text(root, width=44, height=10, foreground="red")
errors.place(x=50, y=100)

root.mainloop()