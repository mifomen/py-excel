import openpyxl

URL = "orders_monitoring_2023_04_13_14_12_20361946.xlsx"
wb = openpyxl.reader.excel.load_workbook(filename=URL, data_only=True)
wb.active = 0
sheetLocal = wb.active

# array.append(atlestCharToInt(sheetLocal['D' + str(i)].value))


def atlestCharToInt (a):
  if a=='':
    return '';
  else:
  #lengthLine = len(a.value)
    s = a[len(a)-3:len(a)].strip()
    return int(s)

gg = []
classLastNames = ['А','Б','В','Г','Д'];

def split(arr, size):
  arrs = []
  while len(arr) > size:
    pice = arr[:size]
    arrs.append(pice)
    arr = arr[size:]
  arrs.append(arr)
  return arrs

for i in classLastNames:
  for j in range (45,50,1):
    s = sheetLocal['B' + str(j)].value;
    print(f"{s} s[len(s)-1] = {s[len(s)-1]}")

    if s[len(s)-1] == i:
      x = atlestCharToInt(sheetLocal['D' + str(j)].value)
      gg.append(x)
      x = atlestCharToInt(sheetLocal['E' + str(j)].value)
      gg.append(x)
      break;

gg=split(gg,2)
print(f"gg = {gg}")