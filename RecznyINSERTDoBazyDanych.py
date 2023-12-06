import pymssql

server = 'dbserver-projekt-zespolowy-dt.database.windows.net'
database = 'db'
username = 'Ladybug'
password = 'FfqGV3PY'

conn = pymssql.connect(server=server, user=username, password=password, database=database)
cursor = conn.cursor()

nowe_dane = {
    'id': '1',
    'rodzaj_alarmu': 'alfa',
    'pododzial': 'ODPion',
    'kto_wprowadzil': 'ODWat'
}

#query = "SELECT username, password FROM users"
query = f"INSERT INTO sygnaly_obsluga ({', '.join(nowe_dane.keys())}) VALUES ({', '.join(['%s'] * len(nowe_dane))})"
values = tuple(nowe_dane.values())
print(query)

cursor.execute(query, values)

conn.commit()

cursor.close()
conn.close()