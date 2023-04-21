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
    majorDimension = DIMENSION  # COLUMNS ROWS
  ).execute()
  return value

spreadsheetId = '1fCLe7lzyYB9NMm8iQxXUzHzeMOE4BP1sYbJJheGV9j4'
listName = 'КУРСЫ'

updatesRangeSheet = 'A2:B3'
majDimension = "ROWS"

ss1 = getValues(spreadsheetId, listName, updatesRangeSheet, majDimension)
# print(ss1)


metaData = []

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

# updatesValues(spreadsheet_id2)

# end time
end = time.time()

print(f"Execution time of the program is- {end-start:5.3f} s.")
# print(f"ss1 = {type(ss1)}")
# print(f"ss1 = {ss1.items()}")
# print(f"ss1 = {ss1.keys()}")
# print(f"ss1 = {ss1.values()}")

metaData = list(ss1.values());
metaData = metaData[2]
# metaData = split(metaData,1)

# print(f"metaData = {metaData[0][0]}") #--right

# print(f"metaData = {metaData}")


import json
file = open("./teachers.json", encoding="utf8")
ny = json.load(file)

print(ny[0]['upQualification'])
# ny[0]['upQualification'] = 'upQualification'
# print(ny[0]['upQualification'])

St = ny[0]['upQualification']
print(f"St = {St}")

# St.replace(", \"[, maxcount])
# St.replace(", \")
# json.dumps(St)
St.replace('"', '\\"') # work this
print(f"St = {St}")


import io
with io.open('data.txt', 'w', encoding='utf-8') as f:
  f.write(json.dumps(St, ensure_ascii=False))


# for i in range(0,len(ny),65):
  # print(ny[i]['upQualification'])
