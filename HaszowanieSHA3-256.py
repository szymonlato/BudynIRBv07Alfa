import hashlib

def haszuj_sha3(tekst):
    # Utwórz obiekt haszujący SHA-3 (SHA3-256)
    sha3_hasz = hashlib.sha3_256()

    # Konwertuj tekst na bajty (UTF-8) i zaktualizuj hasz
    sha3_hasz.update(tekst.encode('utf-8'))

    # Pobierz zhashowany wynik
    zhashowany_wynik = sha3_hasz.hexdigest()

    return zhashowany_wynik

# Przykładowy tekst do haszowania
tekst_do_haszowania = "giI}ky]CddhazoO]"

# Haszuj tekst za pomocą SHA-3
"""
i = 1
while True:
    if i == 4:
        break
    zhashowany_tekst = haszuj_sha3(tekst_do_haszowania + str(i))
    print(f"Tekst do haszowania: {tekst_do_haszowania + str(i)}")
    print(f"Wynik haszowania (SHA3-256): {zhashowany_tekst}")
    i += 1
"""
zhashowany_tekst = haszuj_sha3(tekst_do_haszowania)
print(f"Tekst do haszowania: {tekst_do_haszowania}")
print(f"Wynik haszowania (SHA3-256): {zhashowany_tekst}")

