import sys
from PyQt5.QtWidgets import (QApplication, QWidget, QPushButton, QLabel, QMainWindow, QGridLayout, QTabWidget, QAction,
                             QTableWidget, QVBoxLayout, QInputDialog, QFileDialog, QTableWidgetItem, QTextEdit,
                             QHBoxLayout, QCheckBox)
import shutil
import os
from docx import Document
from database import (create_table_saskaitu_duomenys, create_table_sablonai, insert_data_sablonai, atvaizduoti_sablonus,
                      insert_data_saskaitos, atvaizduoti_saskaitas)

# prisijungimai i duomenu baze
db_params = {
    "host": "localhost",
    "database": "automationApp",
    "user": "postgres",
    "password": "123",
    "port": "5432"
}

class Application(QMainWindow):
    def __init__(self):
        super().__init__()
        self.sablonu_list = []
        self.saskaitu_list = []
        self.sablonu_sarasas = None
        self.saskaitu_sarasas = None
        self.initUI()

    # aplikacijos pagrindinis meniu
    def initUI(self):
        self.showMaximized()
        self.setWindowTitle('Dokumentų Generavimo Sistema')

        # virsutine juosta
        menu_juosta = self.menuBar()

        # pasirinkimai virsutineje juostoje ir susijusiu operaciju mygtukai
        sablonai_menu = menu_juosta.addMenu('Šablonai')
        sablonu_sarasas_button = QAction('Šablonu Sąrašas', self)
        ikelti_sablona_button = QAction('Ikelti Nauja šabloną', self)
        duomenys_menu = menu_juosta.addMenu('Duomenys')
        saskaitu_sarasas_button = QAction('Saskaitų sąrašas', self)
        sutarciu_sarasas_button = QAction('Sutarčių sąrašas', self)

        sablonu_sarasas_button.triggered.connect(self.sablonuSarasas)
        ikelti_sablona_button.triggered.connect(self.ikeltiSablona)
        saskaitu_sarasas_button.triggered.connect(self.saskaituSarasas)

        sablonai_menu.addAction(sablonu_sarasas_button)
        sablonai_menu.addAction(ikelti_sablona_button)
        duomenys_menu.addAction(saskaitu_sarasas_button)
        duomenys_menu.addAction(sutarciu_sarasas_button)

    def sablonuSarasas(self):
        try:

            # jeigu ekrane yra rodomas saskaitu sarasas ji uzdarysime, tam kad rodyti sablonu sarasa
            # tai pat sukuriame jeigu dar neturime duomenu bazeje table kuris laikys musu sablonu duomenis
            # ir pasiemame informacija is db
            if self.saskaitu_sarasas:
                self.saskaitu_sarasas.hide()
            create_table_sablonai(db_params)
            self.sablonu_list = atvaizduoti_sablonus(db_params)
            #print(self.sablonu_list)

            # suformuojame lentele sarasui atvaizduoti
            self.sablonu_sarasas = QTableWidget(self)
            self.sablonu_sarasas.setColumnCount(3)
            self.sablonu_sarasas.setHorizontalHeaderLabels(['Pavadinimas', 'Tipas', ''])

            # priskiriame lentelei tiek eiluciu kiek sablonu sarase yra sablonu
            self.sablonu_sarasas.setRowCount(len(self.sablonu_list))

            # pagal kiekviena eilute ikeliame duomenis i atitinkamus stulpelius
            # sablonas: saraso eilute, stulpelis: kokia butent informacija is saraso eilutes bus ideta lenteles
            # stulpeliuose
            for sablonas, stulpelis in enumerate(self.sablonu_list):
                sablono_name = stulpelis["pavadinimas"]
                sablono_tipas = stulpelis["tipas"]

                # i lentele idedame paimtus saraso duomenis
                self.sablonu_sarasas.setItem(sablonas, 0, QTableWidgetItem(sablono_name))
                self.sablonu_sarasas.setItem(sablonas, 1, QTableWidgetItem(sablono_tipas))

                # kiekvienai eilutei lentele pridedame mygtuka modifikuotis kuris mums leis pasirinkti kuri sablona
                # koreguosime
                self.modifikuoti_button = QPushButton('Modifikuoti', self)
                self.modifikuoti_button.clicked.connect(self.modifikuotiSablona)
                self.sablonu_sarasas.setCellWidget(sablonas, 2, self.modifikuoti_button)

            sablonai_layout = QVBoxLayout()
            sablonai_layout.addWidget(self.sablonu_sarasas)
            central_widget = QWidget(self)
            self.setCentralWidget(central_widget)
            central_widget.setLayout(sablonai_layout)

        except Exception as e:
            print(e)


    def ikeltiSablona(self):
        global naujasis_failo_pavadinimas
        try:
            data = []
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

                        # visus naujus duomenis ikeliame i duomenu baze
                        data.append({"pavadinimas": pavadinimas, "tipas": tipas, "failo_kelias":
                            naujasis_failo_pavadinimas})
                        #print(data)
                        insert_data_sablonai(data,db_params)
            self.sablonuSarasas()
        except Exception as e:
            print(e)

    def modifikuotiSablona(self):
        try:
            # neberodome saraso ekrane
            self.sablonu_sarasas.hide()
            self.fileEdit = QTextEdit(self)

            # gauname indeksa mygtuko kuris buvo paspaustas ir su indeksu pasiemame sablono kelia kuri redaguosime
            index = self.sablonu_sarasas.indexAt(self.modifikuoti_button.pos())
            row = index.row()
            self.modifikuojamas_failas = self.sablonu_list[row]["failo_kelias"]

            # pagal kelia atsidarome faila ir su for pasiemame ir atvaizduojame teksta kuri galesime koreguoti
            dokumentas = Document(self.modifikuojamas_failas)
            tekstas = ""
            for paragraph in dokumentas.paragraphs:
                tekstas += paragraph.text + '\n'

            self.fileEdit.setPlainText(tekstas)

            # buttonu layout
            button_layout = QHBoxLayout()
            issaugoti_button = QPushButton("Issaugoti", self)
            issaugoti_button.clicked.connect(self.atnaujintiSablona)
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
            # atnaujiname faila su naujais irasais, irasa fiksuojame pagal eilute
            dokumentas = Document()
            for eilute in self.fileEdit.toPlainText().split('\n'):
                dokumentas.add_paragraph(eilute)

            dokumentas.save(self.modifikuojamas_failas)
        except Exception as e:
            print(e)

    def saskaituSarasas(self):
        try:
            # toks pat veikimas kaip ir su sablonu sarasu. Isjungiame jeigu rodomas sablonu sarasas, sukuriam table
            # jeigu neturime ir is db pasiemame duomenis ju atvaizdavimui
            if self.sablonu_sarasas:
                self.sablonu_sarasas.hide()
            create_table_saskaitu_duomenys(db_params)
            self.saskaitu_list = atvaizduoti_saskaitas(db_params)
            # print(self.saskaitu_list)

            self.saskaitu_sarasas = QTableWidget(self)
            self.saskaitu_sarasas.setColumnCount(20)
            self.saskaitu_sarasas.setHorizontalHeaderLabels(['Pazymeti', 'Data', 'Serija', 'Numeris', 'Pardavejo imone',
                                                             'Pardavejo adresas', 'Pardavejo kodas',
                                                             'Pardavejo PVM kodas', 'Pirkejo imone', 'Pirkejo adresas',
                                                             'Pirkejo kodas', 'Pirkejo PVM kodas', 'Preke', 'Mat vnt',
                                                             'Kiekis', 'Kaina be PVM', 'Suma be PVM', 'PVM proc',
                                                             'PVM suma', 'Suma'])
            self.saskaitu_sarasas.setRowCount(len(self.saskaitu_list))

            # pirmasis for pasiema visus saskaitos duomenis pagal eilute sarase, o su antruoju for is kiekvieno eilutes
            # langelio pasirenkame informacija is eiles i atvaizduojame ta informacija table
            for saskaita, stulpelis in enumerate(self.saskaitu_list):
                check_box = QCheckBox()
                self.saskaitu_sarasas.setCellWidget(saskaita, 0, check_box)
                for langelis, (kintamasis, verte) in enumerate(stulpelis.items()):
                    self.saskaitu_sarasas.setItem(saskaita, (langelis + 1), QTableWidgetItem(verte))


            button_layout = QHBoxLayout()
            pasirinkti_location_button = QPushButton('Ikelti duomenys')
            atnaujinti_button = QPushButton("Atnaujinti", self)
            pasirinkti_sablona_button = QPushButton("Pasirinkti Šabloną", self)
            formuoti_button = QPushButton("Formuoti Dokumentus", self)
            pasirinkti_location_button.clicked.connect(self.dokumento_pasirinkimas)
            button_layout.addWidget(pasirinkti_location_button)
            button_layout.addWidget(atnaujinti_button)
            button_layout.addWidget(pasirinkti_sablona_button)
            button_layout.addWidget(formuoti_button)

            saskaitu_layout = QVBoxLayout()
            saskaitu_layout.addLayout(button_layout)
            saskaitu_layout.addWidget(self.saskaitu_sarasas)
            central_widget = QWidget(self)
            self.setCentralWidget(central_widget)
            central_widget.setLayout(saskaitu_layout)

        except Exception as e:
            print(e)

    def dokumento_pasirinkimas(self):
        try:
            # pasirenkame dokumenta is kompiuterio (.xlsx) kurio duomenis bus atvaizduojame saskaitu sarase
            failo_kelias, _ = QFileDialog.getOpenFileName(self, 'Pasirinkite Failą', '', 'Visi failai (*)')
            if failo_kelias:
                insert_data_saskaitos(failo_kelias, db_params)
        except Exception as e:
            print(e)


def main():
    app = QApplication(sys.argv)
    ex = Application()
    ex.show()
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()