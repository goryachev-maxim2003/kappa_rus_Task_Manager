import tkinter as tk
import xlwings as xw
import pandas as pd
import numpy as np
import os

files = os.listdir()
answers_file_name = 'MenedzherZadach'
# Ищем файл со строкой 'MenedzherZadach' в названии
for file_name in files:
    # if 'MenedzherZadach' in file_name and file_name[0]!='~':
    #     answers_file_name = file_name
    #     break
    print(file_name)