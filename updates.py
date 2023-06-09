# pip install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib oauth2client
# pip install install google-api-python-client google-auth-httplib2 google-auth-oauthlib oauth2client

import time
# from tkinter.filedialog import askopenfilename
from datetime import datetime
# import openpyxl

import httplib2 # pip install httplib2
import apiclient.discovery
from os import path
from oauth2client.service_account import ServiceAccountCredentials

# CREDENTIALS_FILE = path.expandvars(r'%APPDATA%\Roaming\gspread\service_account.json')

# starting time
start = time.time()


print('Start work')

# авторизация гугл ака, по json файлу
# gc = gspread.service_account()
# открытие гугл таблицы с именем "PythonSheets"
# sh = gc.open("Питание Лицей")
# Выбрать ппервый активный у нее лист
# GoogleSheets = sh.sheet1


def split(arr, size):
    """Функция для разбиения листа на под листы, чтобы отправить на страницу"""
    arrs = []
    while len(arr) > size:
        pice = arr[:size]
        arrs.append(pice)
        arr = arr[size:]
    arrs.append(arr)
    return arrs


CREDENTIALS_FILE = "../service_account.json"



credentials = ServiceAccountCredentials.from_json_keyfile_name(
  CREDENTIALS_FILE,
  ['https://www.googleapis.com/auth/spreadsheets',
   'https://www.googleapis.com/auth/drive'])
httpAuth = credentials.authorize(httplib2.Http())
service = apiclient.discovery.build('sheets', 'v4', http = httpAuth)

# list_name = 'Data'
# import asyncio

def getValues(sheet_ID, listName, RANGE, DIMENSION):
  value = service.spreadsheets().values().get(
    spreadsheetId=sheet_ID,
    range=f"{listName}!{RANGE}",
    majorDimension = DIMENSION  #"COLUMNS" ROWS COLUMNS
  ).execute()
  return value

spreadsheet_id1 = '1zBe_cL1IkwyK9U0zmhFGb5H5m8iWqYhS5PaCmpFNEMU'
spreadsheet_id2 = '1qmqhTdc66yO3SKjMipQ0sIG1cevqY08O0sjSUvLH414'

updatesRangeSheet1 = 'A1:A11'
updatesRangeSheet2 = 'A12:G22'
listName1 = 'Data1'
listName2 = 'Data2'
majDimension = "ROWS"

ss1 = getValues(spreadsheet_id1, listName1, updatesRangeSheet1, majDimension)
print()
ss2 = getValues(spreadsheet_id2, listName2, updatesRangeSheet1, majDimension)

print(f"ss1 = {type(ss1)}")
print(f"ss1 = {ss1.items()}")
print(f"ss1 = {ss1.keys()}")
print(f"ss1 = {ss1.values()}")

metaData = list(ss1.values());
metaData = metaData[2]
# metaData = split(metaData,1)
print(f"metaData = {metaData}")

print()
print(f"ss2 = {ss2}")
print()

batch_update_spreadsheet_request_body  = {
  "valueInputOption": "USER_ENTERED",
  "data": [
    {
      "range": "A13:A24",  # диапазон куда грузим
      "majorDimension": majDimension,
      "values": metaData,  # загружаем обед и полдник у 2АБВГ
    }
  ]
}

def updatesValues(idSheet):
  requestUpdateValues = service.spreadsheets().values().batchUpdate(
    spreadsheetId = idSheet,
    body = batch_update_spreadsheet_request_body
  )
  return requestUpdateValues.execute()

updatesValues(spreadsheet_id2)
# end time
end = time.time()

print(f"Execution time of the program is- {end-start:5.3f} s.")
