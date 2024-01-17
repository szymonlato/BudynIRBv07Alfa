import hashlib

import pymssql

def haszuj_sha3(tekst):
    sha3_hasz = hashlib.sha3_256()
    sha3_hasz.update(tekst.encode('utf-8'))
    zhashowany_wynik = sha3_hasz.hexdigest()
    return zhashowany_wynik

server = 'dbserver-projekt-zespolowy-dt.database.windows.net'
database = 'db'
username = 'Ladybug'
password = 'FfqGV3PY'

loginy = ['Login1', 'Login2']
hasla = ['Haslo1', 'Haslo2']
#aasd
conn = pymssql.connect(server=server, user=username, password=password, database=database)
cursor = conn.cursor()
for i in range(len(loginy)):
    try:
        update_query = f"UPDATE users SET password = %s WHERE username = %s"
        cursor.execute(update_query, (haszuj_sha3(hasla[i]), haszuj_sha3(loginy[i])))
        conn.commit()
    except Exception as e:
        print(f"Błąd aktualizacji danych: {e}")
conn.close()