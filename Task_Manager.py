# pyinstaller --onefile --windowed Task_Manager.py //Для того, чтобы сделать exe файл без консоли

import tkinter as tk
import xlwings as xw
import pandas as pd
import numpy as np
import os
import random
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.options import Options


was_open_all = False
# Открытие файла с ответами
# Поиск файла с ответами в директории
def open_all():
    global text_fault
    global answers
    global problems
    global optional_fields_in_machines
    global plot_fields
    global was_open_all
    global answers_file_name
    global for_Task_Manager
    global files_in_program_dir
    was_open_all = False
    for_Task_Manager = pd.read_excel('Файл для Task_Manager exe.xlsx', 'Параметры', dtype = str, index_col=0)
    files_in_program_dir = os.listdir(for_Task_Manager.loc["Путь к папке с программой", "Значение"])
    answers_file_name = 'Файл из Yandex Forms'
    # Ищем файл со строкой 'MenedzherZadach' в названии
    for file_name in files_in_program_dir:
        if 'MenedzherZadach' in file_name and file_name[0]!='~':
            answers_file_name = file_name
            break
    text_fault	= "Сообщить о проблеме/аномалии/ неисправности"

    answers = pd.read_excel(answers_file_name, index_col="ID")
    # Удаление пустых столбцов
    problems = answers[answers["Выберите задачу"] == text_fault]

    problems["Время создания"] = pd.to_datetime(problems["Время создания"])
    # Изменяем ссылку на фото, чтобы была ссылка на общий яндекс диск
    def convert_link(link):
        if pd.isna(link):
            return np.nan
        else:
            splitted = link.split('%2F')
            return f'{for_Task_Manager.loc["Ссылка на папку Yandex Froms на Yandex диске", "Значение"]}/{splitted[-2]}/{splitted[-1]}'

    problems["Вставьте фотографию"] = problems["Вставьте фотографию"].apply(convert_link)

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
    global journals_books
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
    # Удаляет пустые столбцы, учитывая особый случай
    def remove_unnecessary(row):
        row = row.dropna().reset_index(drop = True)
        #Если у Гофроагрегата выбрана секция другое, то узел не выбирается => нужно добавить и узел = другое. Этот случай определяется
        # тем что строка после удаления пустых значений становится короче на 1 элемент чем должна быть (7 вместо 8)
        if (len(row) == 7):
            row.loc[7] = row.loc[6]
            row.loc[6] = "Другое"
        return row
    plot_problems = plot_problems.apply(remove_unnecessary, axis = 1)
    #Записываем файл
    no_new_data = False #новых данных нет 
    if (len(plot_problems) > 0):
        book = xw.Book(plot_fields.loc[plot, "Журнал"])
        journals_books.append(book)
        sht = book.sheets['Sheet1']
        if (sht.range('A3').value is None):
            first_empty_row = 3
        elif (sht.range('A4').value is None): #end('down') от A3, когда в A4 нет значения работает не корректно
            first_empty_row = 4
        else:
            first_empty_row = sht.range('A3').end('down').row + 1
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
#Возвращает колонку датафрейма или False (нужно для того, чтобы в логических выражениях пустой series ассоциировался с False)
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
    try:
        if (not was_open_all):
            open_all()
        f()
    except Exception as e:
        errors.insert(1.0, str(e)
        +'\n1. Проверьте корректность названия в файле: "Файл для Task_Manager exe.xlsx"\n\
2. Исправьте найденную ошибку и сохраните файл\n\
3. Нажмите кнопку обновить данные или перезагрузите приложение')
def do_nothing():
    pass
def upadte():
    global was_open_all
    was_open_all = False
    execute(do_nothing)
def load_from_yandex():
    global for_Task_Manager
    global download_files
    global answers_file_name
    global files_in_program_dir
    for_Task_Manager = pd.read_excel('Файл для Task_Manager exe.xlsx', 'Параметры', dtype = str, index_col=0)

    my_user_agent =["Mozilla/5.0 (Windows NT 10.0) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36 Edg/124.0.0.0",
    "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36 Edg/124.0.0.0",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36 Edg/124.0.0.0",
    "Mozilla/5.0 (Windows NT 10.0) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36 Edg/123.0.0.0",
    'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36 Edg/123.0.0.0',
    'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36 Edg/123.0.0.0',
    'Mozilla/5.0 (Windows NT 10.0) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36 Edg/122.0.0.0',
    'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36 Edg/122.0.0.0',
    'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36 Edg/122.0.0.0',
    'Mozilla/5.0 (Windows NT 10.0) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36 Edg/121.0.0.0',
    'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36 Edg/121.0.0.0',
    'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36 Edg/121.0.0.0',
    'Mozilla/5.0 (Windows NT 10.0) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36 Edg/120.0.0.0',
    'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36 Edg/120.0.0.0',
    'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36 Edg/120.0.0.0',
    'Mozilla/5.0 (Windows NT 10.0) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36 Edg/119.0.0.0',
    'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36 Edg/119.0.0.0',
    'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36 Edg/119.0.0.0']


    edge_options = Options()
    edge_options.add_argument(f"--user-agent={my_user_agent[random.randint(0,18)]}")
    # edge_options.add_argument('--headless')
    driver = webdriver.Edge(options=edge_options)
    driver.get("https://forms.yandex.ru/admin/")
    time.sleep(random.randint(1,3))

    vhod = driver.find_element(By.XPATH, "/html/body/div[2]/div/div/div[3]/div[2]/a")
    vhod.click()
    time.sleep(random.randint(3,5))

    name_input = driver.find_element(By.ID, "passp-field-login")
    name_input.send_keys(for_Task_Manager.loc["Логин", "Значение"])
    int1 = driver.find_element(By.ID, "passp:sign-in")
    int1.click()
    time.sleep(random.randint(1,3))

    name_input2 = driver.find_element(By.ID, "passp-field-passwd")
    name_input2.send_keys(for_Task_Manager.loc["Пароль", "Значение"])
    int2 = driver.find_element(By.ID, "passp:sign-in")
    int2.click()
    time.sleep(random.randint(3,5))
    form_A = driver.find_element(By.XPATH, f'/html/body/div[3]/div/div[1]/div[2]/div[1]/a[{for_Task_Manager.loc["Номер формы в строке форм Яндекса", "Значение"]}]')
    form_A.click()
    time.sleep(random.randint(1,3))

    otvet_A = driver.find_element(By.XPATH, '/html/body/div[3]/div/div[2]/div[1]/a[5]')
    otvet_A.click()
    time.sleep(random.randint(2,5))

    download = driver.find_element(By.XPATH, '/html/body/div[1]/div/div/article/main/div[1]/div/button[2]')
    download.click()
    time.sleep(10)
    download_files = os.listdir(for_Task_Manager.loc["Путь к папке загрузок", "Значение"])
    answers_file_name = 'Файл из Yandex Forms'
    # Удаляем все файлы с 'MenedzherZadach'
    files_in_program_dir = os.listdir(for_Task_Manager.loc["Путь к папке с программой", "Значение"])
    for file_name in files_in_program_dir:
        if 'MenedzherZadach' in file_name and file_name[0]!='~':
            os.remove(file_name)
    # Ищем файл со строкой 'MenedzherZadach' в названии
    for file_name in download_files:
        if 'MenedzherZadach' in file_name and file_name[0]!='~':
            answers_file_name = file_name
            break
    #Перемещаем скаченный файл в папку с директорией
    # os.rename(os.path.join(for_Task_Manager.loc["Путь к папке загрузок", "Значение"], answers_file_name), 
    #         os.path.join(for_Task_Manager.loc["Путь к папке с программой", "Значение"], answers_file_name))

journals_books = []
root = tk.Tk()
root.title("Заполнение журналов на основе Яндекс форм")
root.geometry("750x500")
bt_load_from_yandex = tk.Button(root, text="Скачать файл из yandex forms", width=50, height=1, command=lambda : load_from_yandex())
bt_load_from_yandex.place(x=50, y=25)
bt_load = tk.Button(root, text="Загрузить новые данные из Скаченного файла", width=50, height=1, command=lambda : execute(load))
bt_load.place(x=50, y=55)
bt_close = tk.Button(root, text="Закрыть все журналы", width=50, height=1, command=lambda : execute(close))
bt_close.place(x=50, y=85)
errors = tk.Text(root, width=80, height=10, foreground="red")
errors.place(x=50, y=130)
update = tk.Button(root, text="Обновить данные", width=20, height=2, command=lambda : upadte())
update.place(x=50, y=305)

root.mainloop()