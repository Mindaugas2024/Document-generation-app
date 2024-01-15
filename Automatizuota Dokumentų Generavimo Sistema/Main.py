import sys
from PyQt5.QtWidgets import (QApplication, QWidget, QPushButton, QLabel, QMainWindow, QGridLayout, QTabWidget, QAction,
                             QTableWidget, QVBoxLayout, QInputDialog, QFileDialog, QTableWidgetItem, QTextEdit,
                             QHBoxLayout)
import shutil
import os
from docx import Document

class Application(QMainWindow):
    def __init__(self):
        super().__init__()
        self.sablonu_list = []
        self.initUI()

    # aplikacijos pagrindinis meniu
    def initUI(self):
        self.showMaximized()
        self.setWindowTitle('Dokumentų Generavimo Sistema')

        # virsutine juosta
        menu_juosta = self.menuBar()

        # pasirinkimai virsutineje juostoje ir susijusiu operaciju mygtukai
        sablonai_menu = menu_juosta.addMenu('Šablonai')
        sablonu_sarasas_button = QAction('Šablonu Sarašas', self)
        ikelti_sablona_button = QAction('Ikelti Nauja šabloną', self)

        sablonu_sarasas_button.triggered.connect(self.sablonuSarasas)
        ikelti_sablona_button.triggered.connect(self.ikeltiSablona)

        sablonai_menu.addAction(sablonu_sarasas_button)
        sablonai_menu.addAction(ikelti_sablona_button)

    def sablonuSarasas(self):
        try:

            # suformuojame lentele sarasui atvaizduoti
            self.sarasas = QTableWidget(self)
            self.sarasas.setColumnCount(3)
            self.sarasas.setHorizontalHeaderLabels(['Pavadinimas', 'Tipas', ''])

            # priskiriame lentelei tiek eiluciu kiek sablonu sarase yra sablonu
            self.sarasas.setRowCount(len(self.sablonu_list))

            # pagal kiekviena eilute ikeliame duomenis i atitinkamus stulpelius
            # sablonas: saraso eilute, stulpelis: kokia butent informacija is saraso eilutes bus ideta lenteles
            # stulpeliuose
            for sablonas, stulpelis in enumerate(self.sablonu_list):
                sablono_name = stulpelis["pavadinimas"]
                sablono_tipas = stulpelis["tipas"]

                # i lentele idedame paimtus saraso duomenis
                self.sarasas.setItem(sablonas, 0, QTableWidgetItem(sablono_name))
                self.sarasas.setItem(sablonas, 1, QTableWidgetItem(sablono_tipas))

                # kiekvienai eilutei lentele pridedame mygtuka modifikuotis kuris mums leis pasirinkti kuri sablona
                # koreguosime
                self.modifikuoti_button = QPushButton('Modifikuoti', self)
                self.modifikuoti_button.clicked.connect(self.modifikuotiSablona)
                self.sarasas.setCellWidget(sablonas, 2, self.modifikuoti_button)

            sablonai_layout = QVBoxLayout()
            sablonai_layout.addWidget(self.sarasas)
            central_widget = QWidget(self)
            self.setCentralWidget(central_widget)
            central_widget.setLayout(sablonai_layout)

        except Exception as e:
            print(e)


    def ikeltiSablona(self):
        global naujasis_failo_pavadinimas
        try:
            dok_tipai = ['Saskaita', 'Sutartis']

            # pirmasis popup langelis kuriame reikia pasirinkti ikeliamo sablono tipa
            tipas, ok = QInputDialog.getItem(self, 'Ikelti naują šabloną', 'Pasirinkite dokumento tipą:', dok_tipai, 0,

                                            False)
            # antrasis popup kuris praso ivesti sablono pavadinima
            if ok and tipas:
                pavadinimas, ok = QInputDialog.getText(self, 'Ikelti naują šabloną', 'Įveskite pavadinimą:')

                # treciasis popup yra pasirinkimas sablono failo kuri norime isikelti
                if ok and pavadinimas:
                    failo_kelias, _ = QFileDialog.getOpenFileName(self, 'Pasirinkite Failą', '', 'Visi failai (*)')
                    if failo_kelias:

                        # direktorija kurioje bus laikomas sablonas
                        issaugojimo_direktorija = ('C:/Users/minde/Documents/GitHub/PYTHON/Automatizuota Dokumentų '
                                                   'Generavimo Sistema/Sablonai')
                        failo_pavadinimas = os.path.basename(failo_kelias)

                        # perkeliame su shutil musu pasirinkta faila i sablonu direktorija
                        shutil.copy(failo_kelias, issaugojimo_direktorija)

                        # pervadiname faila pagal ivestus duomenis popupe
                        keiciamas_failo_pavadinimas = (issaugojimo_direktorija + '/' + failo_pavadinimas)
                        if "pdf" in failo_pavadinimas:
                            naujasis_failo_pavadinimas = (issaugojimo_direktorija + '/' + pavadinimas + ".pdf")
                        if "docx" in failo_pavadinimas:
                            naujasis_failo_pavadinimas = (issaugojimo_direktorija + '/' + pavadinimas + ".docx")
                        os.rename(keiciamas_failo_pavadinimas, naujasis_failo_pavadinimas)
                        # print(naujasis_failo_pavadinimas)

                        # visus ivestus duomenis ir naujojo sablono kelia sukeliame i sablonu sarasa tolimesniam
                        # sablono naudojimui
                        self.sablonu_list.append({"pavadinimas": pavadinimas, "tipas": tipas, "failo_kelias":
                            naujasis_failo_pavadinimas})
                        # print(self.sablonu_list)
        except Exception as e:
            print(e)

    def modifikuotiSablona(self):
        try:
            # neberodome saraso ekrane
            self.sarasas.hide()
            self.fileEdit = QTextEdit(self)

            # gauname indeksa mygtuko kuris buvo paspaustas ir su indeksu pasiemame sablono kelia kuri redaguosime
            index = self.sarasas.indexAt(self.modifikuoti_button.pos())
            row = index.row()
            self.modifikuojamas_failas = self.sablonu_list[row]["failo_kelias"]

            # pagal kelia atsidarome faila ir su for pasiemame ir atvaizduojame teksta kuri galesime koreguoti
            document = Document(self.modifikuojamas_failas)
            tekstas = ""
            for paragraph in document.paragraphs:
                tekstas += paragraph.text + '\n'

            self.fileEdit.setPlainText(tekstas)

            # buttonu layout
            button_layout = QHBoxLayout()
            issaugoti_button = QPushButton("Issaugoti", self)
            atsaukti_button = QPushButton("Atsaukti", self)
            button_layout.addWidget(issaugoti_button)
            button_layout.addWidget(atsaukti_button)

            # pagrindinis layout ir buttonu layout pridejimas
            modifikacijos_layout = QVBoxLayout()
            modifikacijos_layout.addWidget(self.fileEdit)
            modifikacijos_layout.addLayout(button_layout)
            central_widget = QWidget(self)
            self.setCentralWidget(central_widget)
            central_widget.setLayout(modifikacijos_layout)
        except Exception as e:
            print(e)

    def atnaujintiSablona(self):
        try:

        except Exception as e:
        print(e)



def main():
    app = QApplication(sys.argv)
    ex = Application()
    ex.show()
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()