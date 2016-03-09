import ExcelConnection



def findcellinproj (exel_name, sheet_name, main_row, column):
    sheet = ExcelConnection.ExcelConnection('{}'.format(exel_name), '{}'.format(sheet_name))

    if sheet.Cells(main_row, column).value != None:
        data = sheet.Cells(main_row, column).value.strip(' ')
    elif sheet.Cells(main_row, column).value == None:
        try:
            data = 'Пустая ячейка'
        except: data = 0
    return data

