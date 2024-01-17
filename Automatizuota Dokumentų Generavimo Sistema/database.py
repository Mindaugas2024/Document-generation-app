import psycopg2
import openpyxl

def create_table_saskaitu_duomenys(conn_params):
    # susikuriame duomenu baze jeigu dar nebuvo sukurta su visais stulpeliais saskaitoms
    conn = psycopg2.connect(**conn_params)
    cur = conn.cursor()
    create_query = """
        create table if not exists sf_duomenys(
        id serial PRIMARY KEY,
        serija varchar(255),
        numeris varchar(255),
        data varchar(255),
        pardavejo_imone varchar(255),
        pardavejo_adresas varchar(255),
        pardavejo_kodas varchar(255),
        pardavejo_pvm_kodas varchar(255),
        pirkejo_imone varchar(255),
        pirkejo_adresas varchar(255),
        pirkejo_kodas varchar(255),
        pirkejo_pvm_kodas varchar(255),
        preke varchar(255),
        mat_vnt varchar(255),
        kiekis varchar(255),
        kaina_be_pvm varchar(255),
        suma_be_pvm varchar(255),
        pvm_proc varchar(255),
        pvm_suma varchar(255),
        suma varchar(255)
        )
    """
    cur.execute(create_query)
    conn.commit()
    cur.close()
    conn.close()

def read_excel(file_path):
    # pagal gauta failo kelia atsidarome ta excel faila ir sheet ir pasiemame duomenys pagal eilute
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook.active
    data = []
    for row in sheet.iter_rows(values_only=True):
        data.append(row)
    return data

def insert_data_saskaitos(failo_kelias, conn_params):
    data = read_excel(failo_kelias)
    # print(data)
    conn = psycopg2.connect(**conn_params)
    cur = conn.cursor()
    insert_query = """
            insert into sf_duomenys(
            serija,
            numeris,
            data,
            pardavejo_imone,
            pardavejo_adresas,
            pardavejo_kodas,
            pardavejo_pvm_kodas,
            pirkejo_imone,
            pirkejo_adresas,
            pirkejo_kodas,
            pirkejo_pvm_kodas,
            preke,
            mat_vnt,
            kiekis,
            kaina_be_pvm,
            suma_be_pvm,
            pvm_proc,
            pvm_suma,
            suma
            ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
        """
    for row in data[1:]:
        cur.execute(insert_query, row)
    conn.commit()
    cur.close()
    conn.close()

def atvaizduoti_saskaitas(conn_params):
    conn = psycopg2.connect(**conn_params)
    cur = conn.cursor()
    show_query = """
        SELECT * FROM sf_duomenys
    """
    cur.execute(show_query)
    saskaitu_list = []
    for row in cur.fetchall():
        saskaita = {"data": row[1], "serija": row[2], "numeris": row[3], "pardavejo_imone": row[4],
                         "pardavejo_adresas": row[5], "pardavejo_kodas": row[6], "pardavejo_pvm_kodas": row[7],
                         "pirkejo_imone": row[8], "pirkejo_adresas": row[9], "pirkejo_kodas": row[10],
                         "pirkejo_pvm_kodas": row[11], "preke": row[12], "mat_vnt": row[13], "kiekis": row[14],
                         "kaina_be_pvm": row[15], "suma_be_pvm": row[16], "pvm_proc": row[17], "pvm_suma": row[18],
                         "suma": row[19]}
        saskaitu_list.append(saskaita)
    # print(saskaitu_list)
    cur.close()
    conn.close()
    return saskaitu_list


def create_table_sablonai(conn_params):
    conn = psycopg2.connect(**conn_params)
    cur = conn.cursor()
    create_query = """
            create table if not exists sablonai_data(
            id serial PRIMARY KEY,
            pavadinimas varchar(255),
            tipas varchar(255),
            failo_kelias varchar(255)
            )
        """
    cur.execute(create_query)
    conn.commit()
    cur.close()
    conn.close()

def insert_data_sablonai(data, conn_params):
    conn = psycopg2.connect(**conn_params)
    cur = conn.cursor()
    insert_query = """
                insert into sablonai_data(
                pavadinimas,
                tipas,
                failo_kelias
                ) VALUES (%s, %s, %s)
            """

    for i in data:
        cur.execute(insert_query, (i["pavadinimas"], i["tipas"], i["failo_kelias"]))
    conn.commit()
    cur.close()
    conn.close()

def atvaizduoti_sablonus(conn_params):
    conn = psycopg2.connect(**conn_params)
    cur = conn.cursor()
    show_query = """
        SELECT * FROM sablonai_data
    """
    cur.execute(show_query)
    sablonu_list = []
    for row in cur.fetchall():
        sablonas = {"pavadinimas": row[1], "tipas": row[2], "failo_kelias": row[3]}
        sablonu_list.append(sablonas)
    cur.close()
    conn.close()
    return sablonu_list
