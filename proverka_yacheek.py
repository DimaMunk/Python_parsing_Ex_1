import win32com.client

logfile = open('log.txt', 'w+')

Excel = win32com.client.Dispatch("Excel.Application")

try:
    wb = Excel.Workbooks.open(u'C:\\Users\\DimaP\\Downloads\\PYTHON_PROJECTS\\СТГ-568-НС-ОМ_Предписания_01.xlsx')
    logfile.write('Exel is open successfully\n')
except:
    print('Excel is opened\n')
    logfile.write('Error, Excel is already opened\n')

try:
    sheet = wb.sheets('СКЗ')
    logfile.write('Sheet opened successfully\n')
except:
    print('The sheet doesn\'t exist\n')
    logfile.write('Error. The sheet doesn\'t exist\n')

if sheet.Cells(20, 4).MergeCells:
    print('yes')

a = sheet.Cells(20, 4).MergeArea.value
b = len(a)
print(a)
print(b)
