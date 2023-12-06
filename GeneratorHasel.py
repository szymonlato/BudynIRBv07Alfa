import random
import string

def generuj_haslo(dlugosc):
    znaki = string.ascii_letters + string.digits + string.punctuation
    return ''.join(random.choice(znaki) for _ in range(dlugosc))

# Generowanie 16 haseł o długości 16 znaków
ilosc_hasel = 16
dlugosc_hasla = 16

hasla = [generuj_haslo(dlugosc_hasla) for _ in range(ilosc_hasel)]

# Wyświetlenie wygenerowanych haseł
for haslo in hasla:
    print(haslo)