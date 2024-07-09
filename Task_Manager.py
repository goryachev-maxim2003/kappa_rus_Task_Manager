# pyinstaller --onefile --windowed Task_Manager.py //Для того, чтобы сделать exe файл без консоли

import tkinter as tk
import xlwings as xw
import pandas as pd
import numpy as np
import os

was_open_all = False
# Открытие файла с ответами
# Поиск файла с ответами в директории
def open_all():
    global files
    global text_fault
    global answers
    global problems
    global optional_fields_in_machines
    global plot_fields
    global was_open_all
    global answers_file_name
    was_open_all = False
    files = os.listdir()
    answers_file_name = 'MenedzherZadach'
    # Ищем файл со строкой 'MenedzherZadach' в названии
    for file_name in files:
        if 'MenedzherZadach' in file_name and file_name[0]!='~':
            answers_file_name = file_name
            break
    text_fault	= "Сообщить о проблеме/аномалии/ неисправности"

    answers = pd.read_excel(answers_file_name, index_col="ID")
    # Удаление пустых столбцов
    problems = answers[answers["Выберите задачу"] == text_fault]

    problems["Время создания"] = pd.to_datetime(problems["Время создания"])

    # Соответствие между участком и журналом
    optional_fields_in_machines = np.array(["Опишите аномалию", "Вставьте фотографию"])
    # Журналы и необязательные поля участка
    plot_fields = pd.read_excel('Файл для Task_Manager exe.xlsx', 'Названия станков и журналов', dtype = str)
    plot_fields.columns = ["Станок", "Журнал"]
    plot_fields.set_index("Станок", inplace=True)
    # Пока что для всех станков одинаковые необязательные поля
    plot_fields["Необязательные поля"] = plot_fields["Журнал"].apply(lambda x: optional_fields_in_machines)

# Распределяем данные по файлам
def get_max_datatime(ser): #Максимальное время в series (для пустой series возвращает самую старую дату)
    return pd.to_datetime(0) if len(ser) == 0 else ser.max()
# Функция записи определённых ответов в журнал определённого участка
def write_in_plot(plot):
    journal = pd.read_excel(plot_fields.loc[plot, "Журнал"], skiprows=[1])
    # Объединяем дату и время
    journal["datatime"] = journal["Дата"] + journal["Время"].apply(lambda t: pd.to_timedelta(str(t)))
    # Находим проблемы определённого участка
    # plot_problems = problems[(problems["Выберите участок"] == plot) | (problems["Выберите станок"] == plot)]
    plot_problems = problems[(false_if_empty(get_column_or_empty(problems, "Выберите участок") == plot)) | (false_if_empty(get_column_or_empty(problems, "Выберите станок") == plot))]
    # Выбираем новые записи (те которых не было в файле)
    plot_problems = plot_problems[plot_problems["Время создания"] > get_max_datatime(journal["datatime"])]
    #Заполняем прочерками необязательные поля
    for col in plot_fields.loc[plot, "Необязательные поля"]:
        plot_problems.loc[plot_problems[col].isna(), col] = "-"
    #Сортируем по времени
    plot_problems = plot_problems.sort_values("Время создания")
    #Перемещаем ФИО, создаём дату и время
    first_column = plot_problems.pop('ФИО')
    plot_problems.insert(1, "ФИО", first_column)
    plot_problems.insert(0, "Дата", plot_problems["Время создания"].dt.date)
    plot_problems["Время создания"] = plot_problems["Время создания"].dt.time.astype(str)
    #Удаляем лишнюю колонку и лишние пустые значения
    drop_cols = ["Выберите задачу"]
    if (len(get_column_or_empty(problems, "Выберите участок")) > 0):
        drop_cols.append("Выберите участок")
    if (len(get_column_or_empty(problems, "Выберите станок")) > 0):
        drop_cols.append("Выберите станок")
    plot_problems = plot_problems.drop(columns = drop_cols, axis = 1)
    plot_problems = plot_problems.apply(lambda row : row.dropna().reset_index(drop = True), axis = 1)
    #Записываем файл
    no_new_data = False #новых данных нет 
    if (len(plot_problems) > 0):
        book = xw.Book(plot_fields.loc[plot, "Журнал"])
        sht = book.sheets['Sheet1']
        first_empty_row =  3 if (sht.range('A3').value is None) else sht.range('A3').end('down').row + 1
        sht.range(f'A{first_empty_row}').expand(mode='table').value = plot_problems.values
        #Сохраняем файл
        book.save()
    else:
        no_new_data = True
    return no_new_data
#Возвращает колонку датафрейма или пустой Series (нужно для того, чтобы не вызывалась ошибка обращения к несуществующему столбцу)
def get_column_or_empty(df, col): 
    if col in df.columns:
        return df[col]
    else:
        return pd.Series()
#Возвращает колонку датафрейма или пустой False (нужно для того, чтобы в логических выражениях пустой series ассоциировался с False)
def false_if_empty(ser):
    if (len(ser) == 0):
        return False
    else:
        return ser
def load():
    plots = get_column_or_empty(problems, "Выберите участок").dropna().unique()
    plots = np.concatenate((plots[plots != 'Конвертация'], get_column_or_empty(problems, "Выберите станок").dropna().unique()))
    no_new_data = True
    for plot in plots:
        no_new_data *= write_in_plot(plot)
    if (no_new_data):
        errors.insert(1.0, "Новых записей не появилось")
def close():
    for book in journals_books:
        book.close()

def execute(f): #Добавляет проверки перед выполнением функции
    errors.delete('1.0', 'end')
    # try:
    if (not was_open_all):
        open_all()
    f()
#     except Exception as e:
#         errors.insert(1.0, str(e)
#         +'\n1. Проверьте корректность названия в файле: "Файл для Task_Manager exe.xlsx"\n\
# 2. Исправьте найденную ошибку и сохраните файл\n\
# 3. Нажмите кнопку обновить данные или перезагрузите приложение')
def do_nothing():
    pass
def upadte():
    global was_open_all
    was_open_all = False
    execute(do_nothing)

journals_books = []
root = tk.Tk()
root.title("Заполнение журналов на основе Яндекс форм")
root.geometry("750x500")
bt_load = tk.Button(root, text="Загрузить новые данные из yandex forms", width=50, height=1, command=lambda : execute(load))
bt_load.place(x=50, y=25)
bt_close = tk.Button(root, text="Закрыть все журналы", width=50, height=1, command=lambda : execute(close))
bt_close.place(x=50, y=55)
errors = tk.Text(root, width=80, height=10, foreground="red")
errors.place(x=50, y=100)
update = tk.Button(root, text="Обновить данные", width=20, height=2, command=lambda : upadte())
update.place(x=50, y=275)

root.mainloop()