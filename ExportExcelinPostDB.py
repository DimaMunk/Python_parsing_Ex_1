import Make_massiv_of_rows
import ExcelConnection
import psycopg2
import find_cell_in_projects

def exportinpostdb(exel_name,
                   sheet_name,
                   proj_excel_name,
                   proj_list_name,
                   start_row = 4,
                   column_1 = 3,
                   column_2 = 6,
                   column_3 = 7,
                   column_4 = 9,
                   column_5 = 12,
                   column_ustranenie = 10,
                   column_prosrochka = 11):
    print('Connect to DB')
    connect_log = open('C:\\Users\\DimaP\\Downloads\\PYTHON_PROJECTS\\LOGS\\Connect to Postgre DB.txt','w+')

    try:
        conn = psycopg2.connect("dbname = 'VisioDB' user = 'postgres' host = 'localhost' password = '1234'")
        connect_log.write('Base is open')
        print('Base is open')
        connect_log.close()
        cur = conn.cursor()
    except:
        connect_log.write('DB is not open')
        connect_log.close()
    print('Reading Excels...')
    array_of_column_4 = Make_massiv_of_rows.MakeMassiv(exel_name, sheet_name)
    array_of_project_list = Make_massiv_of_rows.MakeMassiv(proj_excel_name, proj_list_name, 2, 1)
    sheet = ExcelConnection.ExcelConnection('{}'.format(exel_name), '{}'.format(sheet_name))
    count = 0
    cells_log = open('C:\\Users\\DimaP\\Downloads\\PYTHON_PROJECTS\\LOGS\\Empty cells in predpis.txt','w+')
    print('Reading Excels success')
    empty_cell_count = 0
    print('Extracting Type of Comment...\n')
    cur.execute('''SELECT ("Типы замечаний".id) FROM PUBLIC."Типы замечаний"''')
    rows = cur.fetchall()
    if len(rows) == 0:

        cur.execute('''INSERT INTO PUBLIC."Типы замечаний" ("id", "Тип замечания") VALUES (1, 'Нарушение технологии строительства');''')

        cur.execute('''INSERT INTO PUBLIC."Типы замечаний" ("id", "Тип замечания") VALUES (2, 'Нарушения связанные с отсутствием
                                                                                          разрешительной документации');''')
        cur.execute('''INSERT INTO PUBLIC."Типы замечаний" ("id", "Тип замечания") VALUES (3, 'Нарушения свяязанные с ведением
                                                                                          исполнительной документации');''')
        cur.execute('''INSERT INTO PUBLIC."Типы замечаний" ("id", "Тип замечания") VALUES (4, 'Нарушения связанные с вовлечением МТР
                                                                                          в производство работ без процедения
                                                                                          процедуры входного контроля,
                                                                                          а также нарушения связанные со
                                                                                          складированием МТР');''')
        cur.execute('''INSERT INTO PUBLIC."Типы замечаний" ("id", "Тип замечания") VALUES (5, 'Нарушения по ОТ, ПБ и ЭБ');''')
        cur.execute('''INSERT INTO PUBLIC."Типы замечаний" ("id", "Тип замечания") VALUES (0, 'Пустая ячейка');''')
        conn.commit()
    print('Extracting Type of Comment success\n')

    print('Extracting Projects...\n')
    cur.execute('''SELECT ("Проекты"."Код проекта")  FROM PUBLIC."Проекты" ''')
    row_project_code = cur.fetchall()
    if len(row_project_code) == 0:
        count_1 = 0
        id_projects = 1
        while count_1 <= len(array_of_project_list) - 1:

            row_array_of_project_list = array_of_project_list[count_1][1]
            while row_array_of_project_list <= array_of_project_list[count_1][1] + array_of_project_list[count_1][2] - 1:
                project_code = find_cell_in_projects.findcellinproj(proj_excel_name, proj_list_name, row_array_of_project_list, 1)
                project_name = find_cell_in_projects.findcellinproj(proj_excel_name, proj_list_name, row_array_of_project_list, 2)
                project_master = find_cell_in_projects.findcellinproj(proj_excel_name, proj_list_name, row_array_of_project_list, 3)
                project_owner = find_cell_in_projects.findcellinproj(proj_excel_name, proj_list_name, row_array_of_project_list, 4)
                project_podryad = find_cell_in_projects.findcellinproj(proj_excel_name, proj_list_name, row_array_of_project_list, 5)
                project_date = find_cell_in_projects.findcellinproj(proj_excel_name, proj_list_name, row_array_of_project_list, 6)
                project_details = find_cell_in_projects.findcellinproj(proj_excel_name, proj_list_name, row_array_of_project_list, 7)
                project_type_of_buildings = find_cell_in_projects.findcellinproj(proj_excel_name, proj_list_name, row_array_of_project_list, 8)
                project_name_of_file = find_cell_in_projects.findcellinproj(proj_excel_name, proj_list_name, row_array_of_project_list, 9)
                project_image = find_cell_in_projects.findcellinproj(proj_excel_name, proj_list_name, row_array_of_project_list, 10)
                project_image_2 = find_cell_in_projects.findcellinproj(proj_excel_name, proj_list_name, row_array_of_project_list, 11)

                cur.execute('''INSERT INTO PUBLIC."Проекты" ("id",
                                                        "Код проекта",
                                                        "Проект",
                                                        "Руководитель",
                                                        "Заказчик",
                                                        "Подрядчики",
                                                        "Период выполнения работ",
                                                        "Детали проекта",
                                                        "Тип строительства",
                                                        "Название папки проекта",
                                                        "Картинка",
                                                        "Карточка проекта") VALUES (
                                                        {}, '{}',
                                                        '{}', '{}',
                                                       '{}', '{}',
                                                        '{}', '{}',
                                                         '{}', '{}',
                                                         '{}','{}');'''.format(id_projects,
                                                                                 project_code,
                                                                                 project_name,
                                                                                 project_master,
                                                                                 project_owner,
                                                                                 project_podryad,
                                                                                 project_date,
                                                                                 project_details,
                                                                                 project_type_of_buildings,
                                                                                 project_name_of_file,
                                                                                 project_image,
                                                                                 project_image_2))
                row_array_of_project_list += 1
                id_projects += 1
                conn.commit()
            count_1 += 1

    print('Extracting Projects success\n')

    print('Extracting Injunctions\n')
    while count <= len(array_of_column_4) - 1:
        main_row = array_of_column_4[count][1]

        while main_row <= array_of_column_4[count][1] + array_of_column_4[count][2] - 1:

                    if sheet.Cells(main_row, column_ustranenie).MergeCells == True:
                        if sheet.Cells(main_row, column_ustranenie).MergeArea.value[0][0] != None:
                            if sheet.Cells(main_row, column_ustranenie).MergeArea.value[0][0] == 'устранено':
                                status_ustranenie = 1
                            elif sheet.Cells(main_row, column_ustranenie).MergeArea.value[0][0] == 'не устранено':
                                status_ustranenie = 1
                        elif sheet.Cells(main_row, column_ustranenie).MergeArea.value[0][0] == None:
                            status_ustranenie = 0
                        if sheet.Cells(main_row, column_ustranenie).MergeArea.value[0][0] == None:
                            cells_log.write('In {}\nRow {} Column {} Is Empty Cell\n\n'.format(str(array_of_column_4[count][0]), main_row, column_ustranenie))
                            empty_cell_count += 1

                    elif sheet.Cells(main_row, column_ustranenie).MergeCells == False:
                        if sheet.Cells(main_row, column_ustranenie).value != None:
                            if sheet.Cells(main_row, column_ustranenie).value == 'устранено':
                                status_ustranenie = 1
                            elif sheet.Cells(main_row, column_ustranenie).value == 'не устранено':
                                status_ustranenie = 0
                        elif sheet.Cells(main_row, column_ustranenie).value == None:
                            status_ustranenie = 0
                        if sheet.Cells(main_row, column_ustranenie).value == None:
                            cells_log.write('In {}\nRow {} Column {} Is Empty Cell\n\n'.format(str(array_of_column_4[count][0]), main_row, column_ustranenie))
                            empty_cell_count += 1


                    if sheet.Cells(main_row, column_prosrochka).MergeCells == True:
                        if sheet.Cells(main_row, column_prosrochka).MergeArea.value[0][0] != None:
                            try:
                                if sheet.Cells(main_row, column_prosrochka).MergeArea.value[0][0] >= 0 :
                                    status_prosrochka = 0
                                elif sheet.Cells(main_row, column_prosrochka).MergeArea.value[0][0] < 0 :
                                    status_prosrochka = 1
                            except: status_prosrochka = 1
                        elif sheet.Cells(main_row, column_prosrochka).MergeArea.value[0][0] == None:
                            status_prosrochka = 1

                        if sheet.Cells(main_row, column_prosrochka).MergeArea.value == None :
                            cells_log.write('In {}\nRow {} Column {} Is Empty Cell\n\n'.format(str(array_of_column_4[count][0]), main_row, column_prosrochka))
                            empty_cell_count += 1

                    elif sheet.Cells(main_row, column_prosrochka).MergeCells == False:
                        if sheet.Cells(main_row, column_prosrochka).value != None:
                            try:
                                if sheet.Cells(main_row, column_prosrochka).value >= 0 :
                                    status_prosrochka = 0
                                elif sheet.Cells(main_row, column_prosrochka).value < 0 :
                                    status_prosrochka = 1
                            except: status_prosrochka = 1
                        elif sheet.Cells(main_row, column_prosrochka).value == 0:
                            status_prosrochka = 1
                        if sheet.Cells(main_row, column_prosrochka).value == None:
                            cells_log.write('In {}\nRow {} Column {} Is Empty Cell\n\n'.format(str(array_of_column_4[count][0]), main_row, column_prosrochka))
                            empty_cell_count += 1

                    control_organ = sheet_name

                    if sheet.Cells(main_row, column_2).MergeCells == True:
                        if sheet.Cells(main_row, column_2).MergeArea.value[0][0] != None:
                            date_of_start = sheet.Cells(main_row, column_2).MergeArea.value[0][0]
                        elif sheet.Cells(main_row, column_2).MergeArea.value[0][0] == None:
                            date_of_start = '2014-01-01 00:00:00+00:00'
                        if sheet.Cells(main_row, column_2).value == 0:
                            cells_log.write('In {}\nRow {} Column {} Is Empty Cell\n\n'.format(str(array_of_column_4[count][0]), main_row, column_2))
                            empty_cell_count += 1

                    elif sheet.Cells(main_row, column_2).MergeCells == False:
                        if sheet.Cells(main_row, column_2).value != None:
                            date_of_start = sheet.Cells(main_row, column_2).value
                        elif sheet.Cells(main_row, column_2).value == None:
                            date_of_start = '2014-01-01 00:00:00+00:00'
                        if sheet.Cells(main_row, column_2).value == None:
                            cells_log.write('In {}\nRow {} Column {} Is Empty Cell\n\n'.format(str(array_of_column_4[count][0]), main_row, column_2))
                            empty_cell_count += 1

                    if sheet.Cells(main_row, column_5).MergeCells == True:
                        if sheet.Cells(main_row, column_5).MergeArea.value[0][0] != None:
                            type_of_zamechanie = sheet.Cells(main_row, column_5).MergeArea.value[0][0]
                        elif sheet.Cells(main_row, column_5).MergeArea.value[0][0] == None:
                            type_of_zamechanie = 0
                        if sheet.Cells(main_row, column_5).MergeArea.value == None:
                            cells_log.write('In {}\nRow {} Column {} Is Empty Cell\n\n'.format(str(array_of_column_4[count][0]), main_row, column_5))
                            empty_cell_count += 1

                    elif sheet.Cells(main_row, column_5).MergeCells == False:
                        if sheet.Cells(main_row, column_5).value != None:
                            type_of_zamechanie = sheet.Cells(main_row, column_5).value
                        elif sheet.Cells(main_row, column_5).value == None:
                            type_of_zamechanie = 0
                        if sheet.Cells(main_row, column_5).value == None:
                            cells_log.write('In {}\nRow {} Column {} Is Empty Cell\n\n'.format(str(array_of_column_4[count][0]), main_row, column_5))
                            empty_cell_count += 1

                    if sheet.Cells(main_row, column_3).MergeCells == True:
                        if sheet.Cells(main_row, column_3).MergeArea.value[0][0] != None:
                            date_of_end = sheet.Cells(main_row, column_3).MergeArea.value[0][0]
                        elif sheet.Cells(main_row, column_3).MergeArea.value[0][0] == None:
                            date_of_end = '2014-01-01 00:00:00+00:00'
                        if sheet.Cells(main_row, column_3).MergeArea.value == None:
                            cells_log.write('In {}\nRow {} Column {} Is Empty Cell\n\n'.format(str(array_of_column_4[count][0]), main_row, column_3))
                            empty_cell_count += 1

                    elif sheet.Cells(main_row, column_3).MergeCells == False:
                        if sheet.Cells(main_row, column_3).value != None:
                            date_of_end = sheet.Cells(main_row, column_3).value
                        elif sheet.Cells(main_row, column_3).value == None:
                            date_of_end = '2014-01-01 00:00:00+00:00'
                        if sheet.Cells(main_row, column_3).value == None:
                            cells_log.write('In {}\nRow {} Column {} Is Empty Cell\n\n'.format(str(array_of_column_4[count][0]), main_row, column_3))
                            empty_cell_count += 1

                    if sheet.Cells(main_row, column_4).MergeCells == True:
                        if sheet.Cells(main_row, column_4).MergeArea.value[0][0] != None:
                            date_of_factend = sheet.Cells(main_row, column_4).MergeArea.value[0][0]
                        elif sheet.Cells(main_row, column_4).MergeArea.value[0][0] == None:
                            date_of_factend = '2014-01-01 00:00:00+00:00'
                        if sheet.Cells(main_row, column_4).MergeArea.value == None:
                            cells_log.write('In {}\nRow {} Column {} Is Empty Cell\n\n'.format(str(array_of_column_4[count][0]), main_row, column_4))
                            empty_cell_count += 1

                    elif sheet.Cells(main_row, column_4).MergeCells == False:
                        if sheet.Cells(main_row, column_4).value != None:
                            date_of_factend = sheet.Cells(main_row, column_4).value
                        elif sheet.Cells(main_row, column_4).value == None:
                            date_of_factend = '2014-01-01 00:00:00+00:00'
                        if sheet.Cells(main_row, column_4).value == None:
                            cells_log.write('In {}\nRow {} Column {} Is Empty Cell\n\n'.format(str(array_of_column_4[count][0]), main_row, column_4))
                            empty_cell_count += 1



####
                    cur.execute('''SELECT ("Подрядчики предписания"."Подрядчик")  FROM PUBLIC."Подрядчики предписания" ''')
                    row_podryadchik = cur.fetchall()
                    sovpedenie_podryadchiki = 0
                    if len(row_podryadchik) == 0:
                        id_podryadchik = 1
                        cur.execute('''INSERT INTO PUBLIC."Подрядчики предписания" ("id", "Подрядчик") VALUES ({}, '{}');'''.format(id_podryadchik, str(array_of_column_4[count][0])))
                        conn.commit()
                    else:
                        for row_1 in row_podryadchik:
                            if array_of_column_4[count][0] == row_1[0]:
                                sovpedenie_podryadchiki += 1
                        if sovpedenie_podryadchiki == 0:
                                cur.execute('''SELECT ("Подрядчики предписания".id)  FROM PUBLIC."Подрядчики предписания" ''')
                                row_id_podryadchik = cur.fetchall()
                                id_podryadchik = max(row_id_podryadchik[-1]) + 1
                                cur.execute('''INSERT INTO PUBLIC."Подрядчики предписания" ("id", "Подрядчик") VALUES ({}, '{}');'''.format(id_podryadchik, array_of_column_4[count][0]))
                                conn.commit()
                        elif sovpedenie_podryadchiki != 0:
                                cur.execute('''SELECT ("Подрядчики предписания".id)  FROM PUBLIC."Подрядчики предписания" WHERE ("Подрядчики предписания"."Подрядчик" = '{}' )'''.format(array_of_column_4[count][0]))
                                row_id_podryadchik = cur.fetchall()
                                id_podryadchik = row_id_podryadchik[0][0]


####
                    cur.execute('''SELECT ("Контролирующие органы"."Контролирующий орган")  FROM PUBLIC."Контролирующие органы" ''')
                    row_control_organ = cur.fetchall()
                    sovpadenie_control_organ = 0

                    if len(row_control_organ) == 0:
                        id_control_organ = 1
                        cur.execute('''INSERT INTO PUBLIC."Контролирующие органы" ("id", "Контролирующий орган") VALUES ({}, '{}');'''.format(id_control_organ, control_organ))
                        conn.commit()
                        conn.rollback
                    else:
                        for row_2 in row_control_organ:
                            if control_organ == row_2[0]:
                                sovpadenie_control_organ += 1
                        if sovpadenie_control_organ == 0:
                                cur.execute('''SELECT ("Контролирующие органы".id)  FROM PUBLIC."Контролирующие органы" ''')
                                row_id_control_organ = cur.fetchall()
                                id_control_organ = max(row_id_control_organ[-1]) + 1
                                cur.execute('''INSERT INTO PUBLIC."Контролирующие органы" ("id", "Контролирующий орган") VALUES ({}, '{}');'''.format(id_control_organ, control_organ))
                                conn.commit()

                        elif sovpadenie_control_organ != 0:
                                cur.execute('''SELECT ("Контролирующие органы".id)  FROM PUBLIC."Контролирующие органы" WHERE ("Контролирующие органы"."Контролирующий орган" = '{}' )'''.format(control_organ))
                                row_id_control_organ = cur.fetchall()
                                id_control_organ = row_id_control_organ[0][0]
####



                    cur.execute('''SELECT ("Предписания".id)  FROM PUBLIC."Предписания" ''')
                    row_predpisania = cur.fetchall()

                    if len(row_predpisania) == 0:
                        id_predpisaniya = 1
                    else:
                        id_predpisaniya = max(row_predpisania[-1]) + 1




                    cur.execute('''SELECT ("Проекты".id)  FROM PUBLIC."Проекты" WHERE ("Проекты"."Код проекта" = '{}' )'''.format(exel_name.strip('_Предписания_01.xlsx')))
                    row_id_proj_code = cur.fetchall()
                    id_projects = row_id_proj_code[0][0]



                    cur.execute('''INSERT INTO public."Предписания" ("id",
                                                  "Контролирующий орган",
                                                  "Подрядчик",
                                                  "Дата выдачи",
                                                  "Плановая дата устранения",
                                                  "Фактическая дата устранения",
                                                  "Тип замечания",
                                                  "Проект",
                                                  "Статус заявки завершение",
                                                  "Статус заявки просрочка") VALUES
                                                  ({}, {}, {}, '{}', '{}', '{}', {}, {}, {}, {});'''.format(id_predpisaniya,
                                                                                    id_control_organ,
                                                                                    id_podryadchik,
                                                                                    date_of_start,
                                                                                    date_of_end,
                                                                                    date_of_factend,
                                                                                    type_of_zamechanie,
                                                                                    id_projects,
                                                                                    status_ustranenie,
                                                                                    status_prosrochka))
                    print('Id {} success\n'.format(id_predpisaniya))
                    conn.commit()
                    main_row += 1
                    id_predpisaniya += 1
        count += 1
    if empty_cell_count == 0:
        cells_log.write('Empty cell are not found')
    else:
        cells_log.write('{} cells are empty'.format(empty_cell_count))
    print('Success')
    print('Empty cells {}'.format(empty_cell_count))
    conn.commit()
    cells_log.close()
    conn.close()



