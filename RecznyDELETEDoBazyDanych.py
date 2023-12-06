import pymssql

server = 'dbserver-projekt-zespolowy-dt.database.windows.net'
database = 'db'
username = 'Ladybug'
password = 'FfqGV3PY'

conn = pymssql.connect(server=server, user=username, password=password, database=database)
cursor = conn.cursor()

pododzial = "ODPion"

try:
    # Przykładowe zapytanie SQL do usunięcia rekordu
    query = f"DELETE FROM sygnaly_obsluga"

    # Wykonaj zapytanie SQL
    cursor.execute(query)
    # Zatwierdź zmiany w bazie danych
    conn.commit()

    print("Rekord został usunięty pomyślnie.")

except Exception as e:
    # W razie błędu wyświetl komunikat
    print(f"Błąd: {e}")
    conn.rollback()

finally:
    cursor.close()
    conn.close()