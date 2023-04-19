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

values = service.spreadsheets().values().get(
  spreadsheetId=spreadsheet_id,
  range='A1:E10',
  majorDimension='ROWS'
).execute()

print(f'values = {values}')
