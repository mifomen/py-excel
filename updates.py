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


def localCellToData(rangeStart, rangeStop, charSheet):
    """Перенос данных из локальной xlsx в память питона"""
    array = []
    for i in range(rangeStart, rangeStop, 1):
        # завтраки 1АБВГД, 2АБВГ 3АБВГ класс
        array.append(atlestCharToInt(sheetLocal[str(charSheet) + str(i)].value))
    # y = split(array, 1)
    # return array
    return split(array, 1)

# Ввод информации в гугл таблицу
# GoogleSheets.update("E1", today)

CREDENTIALS_FILE = "../service_account.json"



credentials = ServiceAccountCredentials.from_json_keyfile_name(
  CREDENTIALS_FILE,
  ['https://www.googleapis.com/auth/spreadsheets',
   'https://www.googleapis.com/auth/drive'])
httpAuth = credentials.authorize(httplib2.Http())
service = apiclient.discovery.build('sheets', 'v4', http = httpAuth)

# list_name = 'Data'
# import asyncio

def getValues(sheet_ID,listName,RANGE, DIMENSION):
  value = service.spreadsheets().values().get(
    spreadsheetId=sheet_ID,
    range=f"{listName}!{RANGE}",
    majorDimension = DIMENSION  #"COLUMNS" ROWS COLUMNS
  ).execute()
  return value

spreadsheet_id1 = '1zBe_cL1IkwyK9U0zmhFGb5H5m8iWqYhS5PaCmpFNEMU'
spreadsheet_id2 = '1qmqhTdc66yO3SKjMipQ0sIG1cevqY08O0sjSUvLH414'

updatesRangeSheet = 'A1:G11'
listName1 = 'Data1'
listName2 = 'Data2'
majDimension = "ROWS"

ss1 = getValues(spreadsheet_id1, listName1, updatesRangeSheet, majDimension)
print()
ss2 = getValues(spreadsheet_id2, listName2, updatesRangeSheet, majDimension)

print(f"ss1 = {ss1}")
print(f"ss2 = {ss2}")


batch_update_spreadsheet_request_body  = {
  "valueInputOption": "USER_ENTERED",
  # "totalUpdatedSheets": 2,
  "data": [
    {
      "range": "Data2!" + updatesRangeSheet,  # диапазон куда грузим
      "majorDimension": majDimension,
      "values": ss1,  # загружаем обед и полдник у 2АБВГ
    }
  ]
}


def updateValuesInGoogleSheets(spreadsheet_id):
  requestUpdateValues = service.spreadsheets().values().batchUpdate(
    spreadsheetId=spreadsheet_id,
    body = batch_update_spreadsheet_request_body
  ).execute()

updateValuesInGoogleSheets(spreadsheet_id2)

# end time
end = time.time()

print(f"Execution time of the program is- {end-start:5.3f} s.")
