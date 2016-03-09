import win32com.client


def ExcelConnection(book_name, sheet_name):
    logfile = open('C:\\Users\\DimaP\\Downloads\\PYTHON_PROJECTS\\LOGS\\log ExcelConnection {}.txt'.format(book_name), 'w+')
    Excel = win32com.client.Dispatch("Excel.Application")

    try:
        wb = Excel.Workbooks.open(u'C:\\Users\\DimaP\\Downloads\\PYTHON_PROJECTS\\{}'.format(book_name))
        logfile.write('Exel {} is open successfully\n'.format(book_name))
    except:
        print('Excel {} is already opened\n'.format(book_name))
        logfile.write('Error, Exel {} is already opened\n'.format(book_name))
    try:
        sheet = wb.sheets('{}'.format(sheet_name))
        logfile.write('Sheet {} is opened successfully\n'.format(sheet_name))
    except:
        print('The sheet {} is doesn\'t exist\n'.format(sheet_name))
        logfile.write('Error. The sheet {} is doesn\'t exist\n'.format(sheet_name))
    logfile.close()

    return sheet

