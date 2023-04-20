# pip install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib oauth2client
# pip install instaa google-api-python-client google-auth-httplib2 google-auth-oauthlib oauth2client

import time
# from tkinter.filedialog import askopenfilename
from datetime import datetime
import openpyxl

import httplib2 # pip install httplib2
import apiclient.discovery
from os import path
from oauth2client.service_account import ServiceAccountCredentials

# CREDENTIALS_FILE = path.expandvars(r'%APPDATA%\Roaming\gspread\service_account.json')

# starting time
start = time.time()

URL = "orders_monitoring_2023_04_18_13_07_20364910.xlsx"

# открываем локальную таблицу менять data_only если надо формулы
wb = openpyxl.reader.excel.load_workbook(filename=URL, data_only=True)
# print(wb.sheetnames) # показать имя листа
wb.active = 0  # активировать самый левый лист в книге
# сохранить в переменную для дальнейшей работы с ним
sheetLocal = wb.active



# функция чтобы найти число в конце строки
def atlestCharToInt(сellValue):
    # если применяем к пустой строке,
    """Извлечь цифры из конца данных ячейки"""
    if (len(сellValue) <= 1) or (сellValue == 0) or (сellValue is None):
        # SS = 0
        return 0  # то вернуть пустую строку
    # lengthLine = len(сellValue.velue)
    # забираем 3 последних символа, оберзаем пробелы
    # SS = сellValue[len(сellValue)-3:len(сellValue)].strip()
    # полученную строку переводим в числовой тип данных
    lengthСell = len(сellValue)
    return int(сellValue[lengthСell-3:lengthСell].strip())

# a = "Питающихся: 25Комплекты:- Завтрак 1-4 класс: 25"
# # СТрока для тестов нахождения последних символов
# s = atlestCharToInt(a); # результат функции для
# нахождения последних цифры в строке


def split(arr, size):
    """Функция для разбиения листа на под листы, чтобы отправить на страницу"""
    arrs = []
    while len(arr) > size:
        pice = arr[:size]
        arrs.append(pice)
        arr = arr[size:]
    arrs.append(arr)
    return arrs


# Строка для инициализации, что начали работать с файлами
print('Start work')

# авторизация гугл ака, по json файлу
# gc = gspread.service_account()
# открытие гугл таблицы с именем "PythonSheets"
# sh = gc.open("Питание Лицей")
# Выбрать ппервый активный у нее лист
# GoogleSheets = sh.sheet1


def localCellToData(rangeStart, rangeStop, charSheet):
    """Перенос данных из локальной xlsx в память питона"""
    array = []
    for i in range(rangeStart, rangeStop, 1):
        # завтраки 1АБВГД, 2АБВГ 3АБВГ класс
        array.append(atlestCharToInt(sheetLocal[str(charSheet) + str(i)].value))
    # y = split(array, 1)
    # return array
    return split(array, 1)


x1 = localCellToData(50, 54, 'D')  # формирование обед у 2АБВГ
x2 = localCellToData(50, 54, 'E')  # формирование полдник у 2АБВГ
x3 = localCellToData(5, 18, 'C')  # завтраки 1АБВГД, 2АБВГ 3АБВГ класс
x4 = localCellToData(18, 22, 'D')  # обеды 4АБВГ класс
x5 = localCellToData(22, 26, 'C')  # завтраки 5АБВГ класс
x6 = localCellToData(26, 33, 'D')  # обеды 6АБВГ 7АБВ класс
x7 = localCellToData(33, 36, 'C')  # завтраки 8АБВ класс
x8 = localCellToData(36, 45, 'C')  # завтраки 9АБВ 10АБВ 11АБВ класс

# print(f"x1 = {x1}")

gg = []  # формирование обед и полдник у 1АБВГД
classLastNames = ['А', 'Б', 'В', 'Г', 'Д']

for i in classLastNames:
    # созданеи нужной последовательности ГПД в 1-х классах
    for j in range(45, 50, 1):
        s = sheetLocal['B' + str(j)].value
        if s[len(s)-1] == i:
            XY = atlestCharToInt(sheetLocal['D' + str(j)].value)
            gg.append(XY)
            XY = atlestCharToInt(sheetLocal['E' + str(j)].value)
            gg.append(XY)
            break

# сохраненный обед и полдник у 1АБВГД
gg = split(gg, 2)

currentDay = datetime.now().day
if currentDay <= 10:
    currentDay = f'0{currentDay}'
currentMonth = datetime.now().month
if currentMonth <= 10:
    currentMonth = f'0{currentMonth}'

currentYear = datetime.now().year
today = f'{currentDay}.{currentMonth}.{currentYear}'

# Ввод информации в гугл таблицу
# GoogleSheets.update("E1", today)

CREDENTIALS_FILE = "../service_account.json"
spreadsheet_id = '18T6n3-Yah8U7TthGsb_GBpEJSWfYyI33416FmJ9WJ0A'


credentials = ServiceAccountCredentials.from_json_keyfile_name(
  CREDENTIALS_FILE,
  ['https://www.googleapis.com/auth/spreadsheets',
   'https://www.googleapis.com/auth/drive'])
httpAuth = credentials.authorize(httplib2.Http())
service = apiclient.discovery.build('sheets', 'v4', http = httpAuth)

# list_name = 'Data'

# import asyncio

def getValues(list_name,RANGE):
  value = service.spreadsheets().values().get(
    spreadsheetId=spreadsheet_id,
    range=f"{list_name}!{RANGE}",
    # str(list_name) + '!' + str(RANGE),
    majorDimension='ROWS'
    ).execute()
  return value

# sss = getValues('Data', 'A1:E10')
# print(f"values = {sss}")

batch_update_spreadsheet_request_body  = {
  "valueInputOption": "USER_ENTERED",
  # "totalUpdatedSheets": 2,
  "data": [
    {
      "range": "Data!C8:C11",  # диапазон куда грузим
      "majorDimension": "ROWS",
      "values": x1,  # загружаем обед и полдник у 2АБВГ
    }, {
      "range": "Data!D8:D11",  # диапазон куда грузим
      "majorDimension": "ROWS",
      "values": x2,  # загружаем обед и полдник у 2АБВГ
    }, {
      "range": "Data!C3:D7",  # диапазон куда грузим
      "majorDimension": "ROWS",
      "values": gg,  # загружаем обед и полдник у 1АБВГД
    }, {
      "range": "Data!G35:G43",  # диапазон куда грузим
      "majorDimension": "ROWS",
      "values": x8,  # загружаем обед и полдник у 1АБВГД
    }, {
      "range": "Data!L3:L15",  # диапазон куда грузим
      "majorDimension": "ROWS",
      "values": x3,  # загружаем завтраки 1АБВГД, 2АБВГ 3АБВГ класс
    }, {
      "range": "M16:M19",  # диапазон куда грузим
      "majorDimension": "ROWS",
      "values": x4,  # загружаем обеды 4АБВГ класс
    }, {
      "range": "G21:G24",  # диапазон куда грузим
      "majorDimension": "ROWS",
      "values": x5,  # загружаем завтраки 5АБВГ класс
    }, {
      "range": "I25:I31",  # диапазон куда грузим
      "majorDimension": "ROWS",
      "values": x6,  # загружаем обеды 6АБВГ 7АБВ класс
    }, {
      "range": "G32:G34",  # диапазон куда грузим
      "majorDimension": "ROWS",
      "values": x7,  # загружаем завтраки 8АБВ класс
    }
  ]
}

requestUpdateValues  = service.spreadsheets().values().batchUpdate(
  spreadsheetId=spreadsheet_id,
  # totalUpdatedSheets=3,
  # sheetId ="1921065082",
  body=batch_update_spreadsheet_request_body
)
responseUpdateValues = requestUpdateValues.execute()


# end time
end = time.time()

print(f"Execution time of the program is- {end-start:5.3f} s.")
