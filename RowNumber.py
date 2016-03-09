import ExcelConnection

def rownumber(exel_name,sheet_name, row = 4, column = 5):
    sheet = ExcelConnection.ExcelConnection(exel_name, sheet_name)
    while sheet.Cells(row, column).value != None:
        row += 1
    ExcelConnection.wb = 0
    return row - 1

