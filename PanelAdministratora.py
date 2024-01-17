import sys

import pymssql
from PyQt5 import uic
from PyQt5.QtCore import QTimer, pyqtSignal
from PyQt5.QtGui import QFont
from PyQt5.QtWidgets import *
from PlikiUi.OknoLogowania_ui import *
from PlikiUi.OknoAdministratora_ui import *
from tkinter import *
import hashlib

server = 'dbserver-projekt-zespolowy-dt.database.windows.net'
database = 'db'
username = 'Ladybug'
password = 'FfqGV3PY'

def haszuj_sha3(tekst):
    sha3_hasz = hashlib.sha3_256()
    sha3_hasz.update(tekst.encode('utf-8'))
    zhashowany_wynik = sha3_hasz.hexdigest()
    return zhashowany_wynik

class Ui_Logowanie(QtWidgets.QDialog):
    def __init__(self):
        super(Ui_Logowanie, self).__init__()
        self.setWindowIcon(QtGui.QIcon("ikona.png"))
        self.initUI()
        self.ui.pushButton.clicked.connect(self.Zaloguj)
        self.ui.pushButton_2.clicked.connect(self.Wyjdz)

    def initUI(self):
        uic.loadUi('PlikiUi/OknoLogowania.ui', self)
        self.ui = Ui_Dialog()
        self.ui.setupUi(self)

    def Zaloguj(self):
        login = self.ui.lineEdit.text()
        haslo = self.ui.lineEdit_2.text()
        conn = pymssql.connect(server=server, user=username, password=password, database=database)
        cursor = conn.cursor()
        query = f"SELECT password FROM users WHERE username = '{haszuj_sha3(login)}'"
        cursor.execute(query)
        rows = cursor.fetchall()
        pobraneHaslo = ''
        for row in rows:
            pobraneHaslo = format(row[0])
            break
        conn.close()
        if pobraneHaslo == haszuj_sha3(haslo):
            letters = ""
            try:
                letters = ''.join(filter(str.isalpha, login))
            except:
                self.showAdditionalWindow("Błędny login!", "Błąd")
            if letters == 'Admin':
                self.close()
                self.showAdditionalWindow(f'Witam Admina kochanego', "Witam :)")
                Ui_PanelAdministratora().exec_()
            else:
                self.showAdditionalWindow("Błędny login!", "Błąd")
        else:
            self.showAdditionalWindow("Błędne hasło!", "Błąd")

    def Wyjdz(self):
        sys.exit(app.exec_())

    def showAdditionalWindow(self, message_text, title_text):
        self.additional_window = QDialog(self)
        self.setWindowIcon(QtGui.QIcon("ikona.png"))
        self.additional_window.setWindowTitle(title_text)
        self.additional_window.setGeometry(200, 200, 300, 100)
        layout = QVBoxLayout()
        label = QLabel(message_text)
        font = QFont()
        font.setPointSize(12)
        label.setFont(font)
        layout.addWidget(label)
        self.additional_window.setLayout(layout)
        timer = QTimer(self)
        timer.timeout.connect(self.closeAdditionalWindow)
        timer.start(30000)
        self.additional_window.exec_()

    def closeAdditionalWindow(self):
        self.additional_window.accept()

class Ui_PanelAdministratora(QtWidgets.QDialog):
    def __init__(self):
        super(Ui_PanelAdministratora, self).__init__()
        self.setWindowIcon(QtGui.QIcon("ikona.png"))
        self.initUI()
        self.cieniowanie()
        self.wypelnijComboBoxy()

    def initUI(self):
        uic.loadUi('PlikiUi/OknoAdministratora.ui', self)
        self.ui = Ui_Dialog4()
        self.ui.setupUi(self)
        self.ui.button_group = QButtonGroup()
        self.ui.button_group.addButton(self.ui.radioButton_4)
        self.ui.button_group.addButton(self.ui.radioButton_5)
        self.ui.button_group.addButton(self.ui.radioButton_10)
        self.ui.button_group.addButton(self.ui.radioButton_7)
        self.ui.button_group.addButton(self.ui.radioButton_6)
        self.ui.button_group.addButton(self.ui.radioButton_8)
        self.ui.button_group.addButton(self.ui.radioButton_9)
        self.ui.button_group2 = QButtonGroup()
        self.ui.button_group2.addButton(self.ui.radioButton)
        self.ui.button_group2.addButton(self.ui.radioButton_2)
        self.ui.button_group2.addButton(self.ui.radioButton_3)
        self.ui.radioButton_4.clicked.connect(self.cieniowanie)
        self.ui.radioButton_5.clicked.connect(self.cieniowanie)
        self.ui.radioButton_6.clicked.connect(self.cieniowanie)
        self.ui.radioButton_7.clicked.connect(self.cieniowanie)
        self.ui.radioButton_8.clicked.connect(self.cieniowanie)
        self.ui.radioButton_9.clicked.connect(self.cieniowanie)
        self.ui.radioButton_10.clicked.connect(self.cieniowanie)
        self.ui.pushButton.clicked.connect(self.wyswietlListe)
        self.ui.pushButton_2.clicked.connect(self.wykonaj)

    def cieniowanie(self):
        self.wycieniujWszystkie()
        if self.ui.radioButton_4.isChecked():
            #Dodaj osobe
            self.ui.comboBox.setEnabled(True)
            self.ui.comboBox_2.setEnabled(True)
            self.ui.lineEdit.setEnabled(True)
            self.ui.label_2.setEnabled(True)
            self.ui.label_3.setEnabled(True)
            self.ui.label_4.setEnabled(True)
        elif self.ui.radioButton_5.isChecked():
            #Usuń osobę
            self.ui.comboBox_3.setEnabled(True)
            self.ui.label_5.setEnabled(True)
        elif self.ui.radioButton_10.isChecked():
            #Edytuj osobę
            self.ui.comboBox_4.setEnabled(True)
            self.ui.comboBox_5.setEnabled(True)
            self.ui.comboBox_6.setEnabled(True)
            self.ui.lineEdit_2.setEnabled(True)
            self.ui.label_10.setEnabled(True)
            self.ui.label_7.setEnabled(True)
            self.ui.label_8.setEnabled(True)
            self.ui.label_9.setEnabled(True)
        elif self.ui.radioButton_7.isChecked():
            #dodaj kompanię
            self.ui.lineEdit_3.setEnabled(True)
            self.ui.comboBox_7.setEnabled(True)
            self.ui.label_6.setEnabled(True)
            self.ui.label_11.setEnabled(True)
        elif self.ui.radioButton_6.isChecked():
            #usuń kompanię
            self.ui.comboBox_8.setEnabled(True)
            self.ui.label_12.setEnabled(True)
        elif self.ui.radioButton_8.isChecked():
            #dodaj batalion
            self.ui.lineEdit_4.setEnabled(True)
            self.ui.label_13.setEnabled(True)
        elif self.ui.radioButton_9.isChecked():
            #usuń batalion
            self.ui.comboBox_9.setEnabled(True)
            self.ui.label_14.setEnabled(True)

    def wycieniujWszystkie(self):
        self.ui.comboBox.setEnabled(False)
        self.ui.comboBox_2.setEnabled(False)
        self.ui.lineEdit.setEnabled(False)
        self.ui.comboBox_3.setEnabled(False)
        self.ui.comboBox_4.setEnabled(False)
        self.ui.comboBox_5.setEnabled(False)
        self.ui.comboBox_6.setEnabled(False)
        self.ui.lineEdit_2.setEnabled(False)
        self.ui.lineEdit_3.setEnabled(False)
        self.ui.comboBox_7.setEnabled(False)
        self.ui.comboBox_8.setEnabled(False)
        self.ui.lineEdit_4.setEnabled(False)
        self.ui.comboBox_9.setEnabled(False)
        self.ui.label_2.setEnabled(False)
        self.ui.label_3.setEnabled(False)
        self.ui.label_4.setEnabled(False)
        self.ui.label_5.setEnabled(False)
        self.ui.label_6.setEnabled(False)
        self.ui.label_7.setEnabled(False)
        self.ui.label_8.setEnabled(False)
        self.ui.label_9.setEnabled(False)
        self.ui.label_10.setEnabled(False)
        self.ui.label_11.setEnabled(False)
        self.ui.label_12.setEnabled(False)
        self.ui.label_13.setEnabled(False)
        self.ui.label_14.setEnabled(False)

    def wypelnijComboBoxy(self):
        conn = pymssql.connect(server=server, user=username, password=password, database=database)
        cursor = conn.cursor()

        query = f"SELECT id FROM Osoba"
        cursor.execute(query)
        rows = cursor.fetchall()

        self.ui.comboBox.clear()
        self.ui.comboBox.addItem("--", 0)
        self.ui.comboBox.addItem("szer.", 1)
        self.ui.comboBox.addItem("st.szer.", 2)
        self.ui.comboBox.addItem("st.szer.spec.", 3)
        self.ui.comboBox.addItem("kpr.", 4)
        self.ui.comboBox.addItem("st.kpr.", 5)
        self.ui.comboBox.addItem("plut.", 6)
        self.ui.comboBox.addItem("sierż.", 7)
        self.ui.comboBox.addItem("st.sierż.", 8)
        self.ui.comboBox.addItem("mł.chor.", 9)
        self.ui.comboBox.addItem("chor.", 10)
        self.ui.comboBox.addItem("st.chor.", 11)
        self.ui.comboBox.addItem("st.chor.sztab.", 12)
        self.ui.comboBox.addItem("ppor.", 13)
        self.ui.comboBox.addItem("por.", 14)
        self.ui.comboBox.addItem("kpt.", 15)
        self.ui.comboBox.addItem("mjr.", 16)
        self.ui.comboBox.addItem("ppłk.", 17)
        self.ui.comboBox.addItem("płk.", 18)
        self.ui.comboBox.addItem("szer.pchor.", 19)
        self.ui.comboBox.addItem("st.szer.pchor.", 20)
        self.ui.comboBox.addItem("st.szer.spec.pchor.", 21)
        self.ui.comboBox.addItem("kpr.pchor.", 22)
        self.ui.comboBox.addItem("st.kpr.pchor.", 23)
        self.ui.comboBox.addItem("plut.pchor.", 24)
        self.ui.comboBox.addItem("sierż.pchor.", 25)

        self.ui.comboBox_5.addItem("--", 0)
        self.ui.comboBox_5.addItem("szer.", 1)
        self.ui.comboBox_5.addItem("st.szer.", 2)
        self.ui.comboBox_5.addItem("st.szer.spec.", 3)
        self.ui.comboBox_5.addItem("kpr.", 4)
        self.ui.comboBox_5.addItem("st.kpr.", 5)
        self.ui.comboBox_5.addItem("plut.", 6)
        self.ui.comboBox_5.addItem("sierż.", 7)
        self.ui.comboBox_5.addItem("st.sierż.", 8)
        self.ui.comboBox_5.addItem("mł.chor.", 9)
        self.ui.comboBox_5.addItem("chor.", 10)
        self.ui.comboBox_5.addItem("st.chor.", 11)
        self.ui.comboBox_5.addItem("st.chor.sztab.", 12)
        self.ui.comboBox_5.addItem("ppor.", 13)
        self.ui.comboBox_5.addItem("por.", 14)
        self.ui.comboBox_5.addItem("kpt.", 15)
        self.ui.comboBox_5.addItem("mjr.", 16)
        self.ui.comboBox_5.addItem("ppłk.", 17)
        self.ui.comboBox_5.addItem("płk.", 18)
        self.ui.comboBox_5.addItem("szer.pchor.", 19)
        self.ui.comboBox_5.addItem("st.szer.pchor.", 20)
        self.ui.comboBox_5.addItem("st.szer.spec.pchor.", 21)
        self.ui.comboBox_5.addItem("kpr.pchor.", 22)
        self.ui.comboBox_5.addItem("st.kpr.pchor.", 23)
        self.ui.comboBox_5.addItem("plut.pchor.", 24)
        self.ui.comboBox_5.addItem("sierż.pchor.", 25)

        self.ui.comboBox_3.clear()
        self.ui.comboBox_4.clear()
        for row in rows:
            self.ui.comboBox_3.addItem(format(row[0]), int(format(row[0])) - 1)
            self.ui.comboBox_4.addItem(format(row[0]), int(format(row[0])) - 1)

        query = f"SELECT * FROM Batalion"
        cursor.execute(query)
        rows = cursor.fetchall()

        self.ui.comboBox_7.clear()
        self.ui.comboBox_9.clear()
        for row in rows:
            self.ui.comboBox_7.addItem(format(row[1]), int(format(row[0])) - 1)
            self.ui.comboBox_9.addItem(format(row[0]), int(format(row[0])) - 1)

        query = f"SELECT id, numer_kompanii FROM Kompania"
        cursor.execute(query)
        rows = cursor.fetchall()

        self.ui.comboBox_2.clear()
        self.ui.comboBox_6.clear()
        self.ui.comboBox_8.clear()
        for row in rows:
            self.ui.comboBox_2.addItem(format(row[1]), int(format(row[0])) - 1)
            self.ui.comboBox_6.addItem(format(row[1]), int(format(row[0])) - 1)
            self.ui.comboBox_8.addItem(format(row[0]), int(format(row[0])) - 1)

    def wyswietlListe(self):
        conn = pymssql.connect(server=server, user=username, password=password, database=database)
        cursor = conn.cursor()
        self.ui.textBrowser.clear()
        if self.ui.radioButton.isChecked():
            #Lista osób
            query = f"SELECT * FROM Osoba"
            cursor.execute(query)
            rows = cursor.fetchall()
            linia = f"id   stopień   imie_i_nazwisko     powód   id_kompani\n"
            self.ui.textBrowser.insertHtml(f'<pre><b>{linia}</b></pre>')
            linia = f"-----------------------------------------------------\n"
            self.ui.textBrowser.insertHtml(f'<pre><b>{linia}</b></pre>')
            for row in rows:
                linia = ""
                spaces = "&nbsp;" * (5 - len(format(row[0])))
                linia += f"{format(row[0])}{spaces}"
                spaces = "&nbsp;" * (10 - len(format(row[1])))
                linia += f"{format(row[1])}{spaces}"
                spaces = "&nbsp;" * (20 - len(format(row[2])))
                linia += f"{format(row[2])}{spaces}"
                spaces = "&nbsp;" * (8 - len(format(row[4])))
                linia += f"{format(row[4])}{spaces}"
                spaces = "&nbsp;" * (5 - len(format(row[5])))
                linia += f"{format(row[5])}{spaces}\n"
                self.ui.textBrowser.insertHtml(f'<pre><b>{linia}</b></pre>')
        elif self.ui.radioButton_2.isChecked():
            #Lista kompanii
            query = f"SELECT * FROM Kompania"
            cursor.execute(query)
            rows = cursor.fetchall()
            linia = f"id   nazwa_kompani  id_batalionu\n"
            self.ui.textBrowser.insertHtml(f'<pre><b>{linia}</b></pre>')
            linia = f"-----------------------------------------------------\n"
            self.ui.textBrowser.insertHtml(f'<pre><b>{linia}</b></pre>')
            for row in rows:
                linia = ""
                spaces = "&nbsp;" * (5 - len(format(row[0])))
                linia += f"{format(row[0])}{spaces}"
                spaces = "&nbsp;" * (15 - len(format(row[1])))
                linia += f"{format(row[1])}{spaces}"
                spaces = "&nbsp;" * (5 - len(format(row[2])))
                linia += f"{format(row[2])}{spaces}\n"
                self.ui.textBrowser.insertHtml(f'<pre><b>{linia}</b></pre>')
        elif self.ui.radioButton_3.isChecked():
            #Lista batalionów
            query = f"SELECT * FROM Batalion"
            cursor.execute(query)
            rows = cursor.fetchall()
            linia = f"id   nazwa_batalionu\n"
            self.ui.textBrowser.insertHtml(f'<pre><b>{linia}</b></pre>')
            linia = f"-----------------------------------------------------\n"
            self.ui.textBrowser.insertHtml(f'<pre><b>{linia}</b></pre>')
            for row in rows:
                linia = ""
                spaces = "&nbsp;" * (5 - len(format(row[0])))
                linia += f"{format(row[0])}{spaces}"
                spaces = "&nbsp;" * (15 - len(format(row[1])))
                linia += f"{format(row[1])}{spaces}\n"
                self.ui.textBrowser.insertHtml(f'<pre><b>{linia}</b></pre>')

    def wykonaj(self):
        conn = pymssql.connect(server=server, user=username, password=password, database=database)
        cursor = conn.cursor()

        if self.ui.radioButton_4.isChecked():
            # Dodaj osobe
            query = f"SELECT * FROM Kompania"
            cursor.execute(query)
            rows = cursor.fetchall()
            nowe_dane = {
                'id': self.ui.comboBox_3.count() + 1,
                'stopien': self.ui.comboBox.currentText(),
                'imie_nazwisko': self.ui.lineEdit.text(),
                'stan': 'obecny',
                'powod': '',
                'kompania_id': ''
            }
            for row in rows:
                if format(row[1]) == self.ui.comboBox_2.currentText():
                    nowe_dane['kompania_id'] = format(row[0])
                    break
            query = f"INSERT INTO Osoba ({', '.join(nowe_dane.keys())}) VALUES ({', '.join(['%s'] * len(nowe_dane))})"
            values = tuple(nowe_dane.values())
            cursor.execute(query, values)
            conn.commit()
            cursor.close()
            conn.close()
            WyswietlenieKomunikatu(f"Dodałem Osobę o podanych wartościach:\n {nowe_dane['id']} {nowe_dane['stopien']} {nowe_dane['imie_nazwisko']} {nowe_dane['kompania_id']}")
        elif self.ui.radioButton_5.isChecked():
            try:
                query = f"DELETE FROM Osoba WHERE id={self.ui.comboBox_3.currentText()}"
                cursor.execute(query)
                conn.commit()
                WyswietlenieKomunikatu(f"Usunąłem Osobę o id {self.ui.comboBox_3.currentText()}")
            except Exception as e:
                print(f"Błąd: {e}")
                conn.rollback()
            finally:
                cursor.close()
                conn.close()
        elif self.ui.radioButton_10.isChecked():
            query = f"SELECT * FROM Kompania"
            cursor.execute(query)
            rows = cursor.fetchall()
            noweID = 1
            for row in rows:
                if format(row[1]) == self.ui.comboBox_6.currentText():
                    noweID = format(row[0])
                    break
            try:
                zapytanie = f"UPDATE Osoba SET stopien='{self.ui.comboBox_5.currentText()}', kompania_id={noweID}, imie_nazwisko='{self.ui.lineEdit_2.text()}' WHERE id={self.ui.comboBox_4.currentText()}"
                cursor.execute(zapytanie)
                conn.commit()
            except Exception as e:
                print(f'Błąd: {e}')
            finally:
                conn.close()
            WyswietlenieKomunikatu(f"Edytowałem Osobe o id {self.ui.comboBox_4.currentText()}")
        elif self.ui.radioButton_7.isChecked():
            query = f"SELECT * FROM Batalion"
            cursor.execute(query)
            rows = cursor.fetchall()
            nowe_dane = {
                'id': self.ui.comboBox_8.count() + 1,
                'numer_kompanii': self.ui.lineEdit_3.text(),
                'batalion_id': ''
            }
            for row in rows:
                if format(row[1]) == self.ui.comboBox_7.currentText():
                    nowe_dane['batalion_id'] = format(row[0])
                    break
            query = f"INSERT INTO Kompania ({', '.join(nowe_dane.keys())}) VALUES ({', '.join(['%s'] * len(nowe_dane))})"
            values = tuple(nowe_dane.values())
            cursor.execute(query, values)
            conn.commit()
            cursor.close()
            conn.close()
            WyswietlenieKomunikatu(f"Dodałem Kompanię o podanych wartościach:\n {nowe_dane['id']} {nowe_dane['numer_kompanii']} {nowe_dane['batalion_id']}")
        elif self.ui.radioButton_6.isChecked():
            try:
                query = f"DELETE FROM Kompania WHERE id={self.ui.comboBox_8.currentText()}"
                cursor.execute(query)
                conn.commit()
                WyswietlenieKomunikatu(f"Usunąłem Kompanię o id {self.ui.comboBox_8.currentText()}")
            except Exception as e:
                print(f"Błąd: {e}")
                conn.rollback()
            finally:
                cursor.close()
                conn.close()
        elif self.ui.radioButton_8.isChecked():
            query = f"INSERT INTO Batalion (id, numer_batalionu) VALUES {self.ui.comboBox_9.count() + 1, self.ui.lineEdit_4.text()}"
            cursor.execute(query)
            conn.commit()
            cursor.close()
            conn.close()
            WyswietlenieKomunikatu(f"Dodałem Batalion o podanych wartościach:\n {self.ui.comboBox_9.count() + 1}, {self.ui.lineEdit_4.text()}")
        elif self.ui.radioButton_9.isChecked():
            try:
                query = f"DELETE FROM Batalion WHERE id={self.ui.comboBox_9.currentText()}"
                cursor.execute(query)
                conn.commit()
                WyswietlenieKomunikatu(f"Usunąłem Batalion o id {self.ui.comboBox_9.currentText()}")
            except Exception as e:
                print(f"Błąd: {e}")
                conn.rollback()
            finally:
                cursor.close()
                conn.close()

        self.wypelnijComboBoxy()

def WyswietlenieKomunikatu(tekst):
    oknoKomunikatu = Tk()
    oknoKomunikatu.title("Komunikat")
    x = (oknoKomunikatu.winfo_screenwidth() - 600) // 2
    y = (oknoKomunikatu.winfo_screenheight() - 60) // 2
    oknoKomunikatu.geometry('{}x{}+{}+{}'.format(600, 60, x, y))
    etykieta = Label(oknoKomunikatu, text=f"{tekst}", justify=CENTER, font=30, fg="red")
    etykieta.pack()
    oknoKomunikatu.mainloop()

app = QApplication(sys.argv)
ex = Ui_Logowanie()
ex.show()
sys.exit(app.exec_())