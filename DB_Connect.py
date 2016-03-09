import psycopg2


con = psycopg2.connect("dbname = 'VisioDB' user = 'postgres' host = 'localhost' password = '1234'")
cur = con.cursor()
print('Base open')




cur.execute('''SELECT ("Подрядчики предписания".id)  FROM PUBLIC."Подрядчики предписания" WHERE ("Подрядчики предписания"."Подрядчик" = 'ООО "ПБР" ' )''')
rows = cur.fetchall()

print('{}'.format(rows))
#conn.commit()
con.close()