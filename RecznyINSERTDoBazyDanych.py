import hashlib

import pymssql

server = 'dbserver-projekt-zespolowy-dt.database.windows.net'
database = 'db'
username = 'Ladybug'
password = 'FfqGV3PY'

conn = pymssql.connect(server=server, user=username, password=password, database=database)
cursor = conn.cursor()

def haszuj_sha3(tekst):
    sha3_hasz = hashlib.sha3_256()
    sha3_hasz.update(tekst.encode('utf-8'))
    zhashowany_wynik = sha3_hasz.hexdigest()
    return zhashowany_wynik

nowe_dane = {
    'id': '31',
    'username': haszuj_sha3('Dowodca12'),
    'password': haszuj_sha3('#S)ROw{Rbymn?[a#'),
}

#query = "SELECT username, password FROM users"
query = f"INSERT INTO users ({', '.join(nowe_dane.keys())}) VALUES ({', '.join(['%s'] * len(nowe_dane))})"
values = tuple(nowe_dane.values())
print(query)

cursor.execute(query, values)

conn.commit()

cursor.close()
conn.close()