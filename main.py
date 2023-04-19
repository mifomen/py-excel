
# %appdata%  Roaming\gspread\service_account.json
import time

import pip
pip.main(["install", "gspread"])
pip.main(["install", "openpyxl"])

import gspread
import openpyxl  # pip install openpyxl gspread


# from imp import find_module

# def checkPythonmod(mod):
#     try:
#         op = find_module(mod)
#         return True
#     except ImportError:
#         return False
# try:
#   import openpyxl
# except ImportError as e:
#   pip.main(["install", "openpyxl"])
#   import openpyxl

# try:
#   import gspread
# except ImportError as e:
#   pip.main(["install", "gspread"])
#   import gspread

# if python -c 'import pkgutil; exit(not pkgutil.find_loader("pandas"))'; then
#     echo 'pandas found'
# else
#     echo 'pandas not found'
# fi

from datetime import datetime
# starting time
start = time.time()


# URL = "orders_monitoring_2023_04_18_13_07_20364910.xlsx"

from tkinter.filedialog import askopenfilename #запрос на указание файла с данными
URL = askopenfilename() # запрос на указание файла с данными

# FILE_NAME = 'test.xlsx'
# wb = openpyxl.reader.excel.load_workbook(filename=FILE_NAME,data_only=True)

# открываем локальную таблицу менять data_only если надо формулы
wb = openpyxl.reader.excel.load_workbook(filename=URL, data_only=True)
# print(wb.sheetnames) # показать имя листа
wb.active = 0  # активировать самый левый лист в книге
sheetLocal = wb.active  # сохранить в переменную для дальнейшей работы с ним


def atlestCharToInt (a): # функция чтобы найти число в конце строки
  if (a == '') or (a == 0) or ( a is None ) or (len(a) <= 2):# если применяем к пустой строке,
    s = 0
    return s  # то вернуть пустую строку
  else:
  #lengthLine = len(a.value)
    s = a[len(a)-3:len(a)].strip() # забираем 3 последних символа, оберзаем пробелы
    return int(s) # полученную строку переводим в числовой тип данных

# a = "Питающихся: 25Комплекты:- Завтрак 1-4 класс: 25"   # СТрока для тестов нахождения последних символов
# s = atlestCharToInt(a); # результат функции для нахождения последних цифры в строке

def split(arr, size): # функция для разбиения листа на под листы, чтбоы отправить на страницу
  arrs = []
  while len(arr) > size:
    pice = arr[:size]
    arrs.append(pice)
    arr = arr[size:]
  arrs.append(arr)
  return arrs

print('Start work') #Строка для инициализации, что начали работать с файлами

gc = gspread.service_account() #авторизация гугл ака, по json файлу
sh = gc.open("Питание Лицей") #открытие гугл таблицы с именем "PythonSheets"
GoogleSheets = sh.sheet1 #Выбрать ппервый активный у нее лист

array = []
for i in range(50,54,1): # формирование обед и полдник у 2АБВГ
  array.append(atlestCharToInt(sheetLocal['D' + str(i)].value))
  array.append(atlestCharToInt(sheetLocal['E' + str(i)].value))
x = split(array,2) # пакуем в подлисты по 2 для вставки в гугл таблицу

array = []
for i in range(5,18,1): #завтраки 1АБВГД, 2АБВГ 3АБВГ класс
  array.append(atlestCharToInt(sheetLocal['C' + str(i)].value))
y = split(array,1)

array = []
for i in range(18,22,1): #обеды 4АБВГ класс
  array.append(atlestCharToInt(sheetLocal['D' + str(i)].value))
z = split(array,1)

array = []
for i in range(22,26,1): #завтраки 5АБВГ класс
  array.append(atlestCharToInt(sheetLocal['C' + str(i)].value))
v = split(array,1)

array = []
for i in range(26,33,1): # обеды 6АБВГ 7АБВ класс
  array.append(atlestCharToInt(sheetLocal['D' + str(i)].value))
b = split(array,1)

array = []
for i in range(33,36,1): #завтраки 8АБВ класс
  array.append(atlestCharToInt(sheetLocal['C' + str(i)].value))
n = split(array,1)

array = []
for i in range(36,45,1): #завтраки 9АБВ 10АБВ 11АБВ класс
  array.append(atlestCharToInt(sheetLocal['C' + str(i)].value))
nn9 = split(array,1)


gg = [] # формирование обед и полдник у 1АБВГД
classLastNames = ['А','Б','В','Г','Д'];

for i in classLastNames: #созданеи нужной последовательности ГПД в 1-х классах
  for j in range (45,50,1):
    s = sheetLocal['B' + str(j)].value;
    if s[len(s)-1] == i:
      xy = atlestCharToInt(sheetLocal['D' + str(j)].value)
      gg.append(xy)
      xy = atlestCharToInt(sheetLocal['E' + str(j)].value)
      gg.append(xy)
      break;

gg = split(gg,2)
# print(f"gg = {gg}")

currentDay = datetime.now().day
if currentDay <=10:
  currentDay = f'0{currentDay}'
currentMonth = datetime.now().month
if currentMonth <= 10:
  currentMonth = f'0{currentMonth}'

currentYear = datetime.now().year
today = f'{currentDay}.{currentMonth}.{currentYear}'
#
# today = "=TODAY()"

# print('today ' + str(today))
# today = split(today,1)
#Ввод информации в гугл таблицу

GoogleSheets.update("E1", today)
GoogleSheets.batch_update([{
    'range': 'C8:D11',  # диапазон куда грузим
    'values': x,  #  загружаем обед и полдник у 2АБВГ
},{
  'range': 'C3:D7',  # диапазон куда грузим
  'values': gg,  # загружаем обед и полдник у 1АБВГД
},{
  'range': 'G35:G43',  # диапазон куда грузим
  'values': nn9,  # загружаем обед и полдник у 1АБВГД
},{
  'range': 'L3:L15',  # диапазон куда грузим
    'values': y, #загружаем завтраки 1АБВГД, 2АБВГ 3АБВГ класс
},{
  'range': 'M16:M19',  # диапазон куда грузим
  'values': z,#загружаем обеды 4АБВГ класс
},{
  'range': 'G21:G24',  # диапазон куда грузим
  'values': v,  #загружаем завтраки 5АБВГ класс
},{
  'range': 'I25:I31',  # диапазон куда грузим
  'values': b,  # загружаем обеды 6АБВГ 7АБВ класс
},{
  'range': 'G32:G34',  # диапазон куда грузим
  'values': n,  #загружаем завтраки 8АБВ класс
}])

# end time
end = time.time()

print(f"Execution time of the program is- {end-start:5.3f} s.")
