import ExcelConnection
import RowNumber

def MakeMassiv(exel_name,sheet_name, row_number = 4, column_number = 4):
    array = []
    counter = row_number
    rows = RowNumber.rownumber(exel_name, sheet_name, row_number, column_number + 1)
    sheet = ExcelConnection.ExcelConnection(exel_name, sheet_name)

    while counter <= rows:
        alist = []
        if sheet.Cells(counter, column_number).MergeCells == True:
            if sheet.Cells(counter, column_number).MergeArea.value[0] != None:
                alist.append(sheet.Cells(counter, column_number).MergeArea.value[0][0].strip(' '))
            else:
                alist.append(sheet.Cells(counter, column_number).MergeArea.value[0])
                log_file = open('C:\\Users\\DimaP\\Downloads\\PYTHON_PROJECTS\\LOGS\\Error File {} stroka {} colonka {}.txt'.format(exel_name, row_number, row_number),'w+')
                log_file.write('row {} is None'.format(counter))
                log_file.close()
            alist.append(counter)
            alist.append(len(sheet.Cells(counter, column_number).MergeArea.value))
            array.append(alist)
            counter += len(sheet.Cells(counter, column_number).MergeArea.value)

        elif sheet.Cells(counter, column_number).MergeCells == False:
            if sheet.Cells(counter, column_number).value != None:
                alist.append(sheet.Cells(counter, column_number).value.strip(' '))
            else:
                alist.append(sheet.Cells(counter, column_number).value)
                log_file = open('C:\\Users\\DimaP\\Downloads\\PYTHON_PROJECTS\\LOGS\\Error File {} stroka {} colonka {}.txt'.format(exel_name, row_number, row_number),'w+')
                log_file.write('row {} is None'.format(counter))
                log_file.close()
            alist.append(counter)
            alist.append(1)
            array.append(alist)
            counter += 1
    return array

