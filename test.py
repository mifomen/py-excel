import httplib2 # pip install httplib2
import apiclient.discovery
from os import path
from oauth2client.service_account import ServiceAccountCredentials

# CREDENTIALS_FILE = path.expandvars(r'%APPDATA%\Roaming\gspread\service_account.json')

CREDENTIALS_FILE = "../service_account.json"
spreadsheet_id = '18T6n3-Yah8U7TthGsb_GBpEJSWfYyI33416FmJ9WJ0A'


credentials = ServiceAccountCredentials.from_json_keyfile_name(
  CREDENTIALS_FILE,
  ['https://www.googleapis.com/auth/spreadsheets',
   'https://www.googleapis.com/auth/drive'])
httpAuth = credentials.authorize(httplib2.Http())
service = apiclient.discovery.build('sheets', 'v4', http = httpAuth)

list_name = 'Data'

values = service.spreadsheets().values().get(
  spreadsheetId=spreadsheet_id,
  range=list_name + '!' + 'A1:E10',
  majorDimension='ROWS'
).execute()

print(f'values = {values}')

values = service.spreadsheets().values().batchUpdate(
  spreadsheetId=spreadsheet_id,
  body={
    "valueInputOption": "USER_ENTERED",
    "data": [
      {
       "range": "B3:C4",
       "majorDimension": "ROWS",
       "values": [["this b3", "this c3"], ["this b4", "this c4"]]
        },
      {
        "range": "D5:E6",
        "majorDimension": "COLUMNS",
        "values": [["this d5", "this d6"], ["this e5", "=5+5"]]
      }
    ]
  }
).execute()
