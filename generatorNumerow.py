import random
import hashlib

def generuj_i_zapisz_liste_osob(ilosc_osob, klucz, nazwa_pliku="Dokumenty/Spis_Numerow.txt"):
    with open(nazwa_pliku, "w") as plik:
        indeksKompani = 1
        indeksOsoby = 1
        for i in range(1, ilosc_osob + 1):
            if indeksOsoby == 6:
                indeksOsoby = 1
                indeksKompani += 1
            seed = int(hashlib.md5(f"{klucz}-{i}".encode()).hexdigest(), 16)
            random.seed(seed)
            numer_telefonu = f"{random.randint(100, 999)}-{random.randint(100, 999)}-{random.randint(100, 999)}"
            osoba = f"Osoba{indeksOsoby}_{indeksKompani}: {numer_telefonu}\n"
            indeksOsoby += 1
            plik.write(osoba)

# Przykład użycia funkcji do wygenerowania i zapisania listy 19 osób do pliku "lista_osob.txt"
generuj_i_zapisz_liste_osob(60, "ABC")