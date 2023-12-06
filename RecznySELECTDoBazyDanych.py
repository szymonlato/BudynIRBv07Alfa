import pymssql

server = 'dbserver-projekt-zespolowy-dt.database.windows.net'
database = 'db'
username = 'Ladybug'
password = 'FfqGV3PY'

conn = pymssql.connect(server=server, user=username, password=password, database=database)
cursor = conn.cursor()

query = f"SELECT * FROM sygnaly_obsluga"
cursor.execute(query)
rows = cursor.fetchall()

for row in rows:
    print(format(row))


#Osoba3_9
#Osoba3_6
