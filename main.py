import xlwings
from xlwings import *
import threading
import tkinter as tk
from tkinter import *
from tkinter import filedialog
from threading import Thread
import mysql.connector
from mysql.connector import Error
from datetime import datetime

arquivogetnet = None
pastadetrabalhogetnet = None
valoresgetnet = []
datasgetnet = []

arquivocbb = None
pastadetrabalhocbb = None
valorescbb = []
datascbb = []

datastr = []

results = []


def create_server_connection(host_name, user_name, user_password):
    connection = None
    try:
        connection = mysql.connector.connect(
            host=host_name,
            user=user_name,
            passwd=user_password
        )
        print("MySQL Database connection successful")
    except Error as err:
        print(f"Error: '{err}'")
    return connection

def create_database(connection, query):
    cursor = connection.cursor()
    try:
        cursor.execute(query)
        print("Database created successfully")
    except Error as err:
        print(f"Error: '{err}'")

def execute_query(connection, query):
    cursor = connection.cursor()
    try:
        cursor.execute(query)
        connection.commit()
        print("Query successful")
    except Error as err:
        print(f"Error: '{err}'")

def ler1():
    global arquivogetnet
    arquivogetnet = filedialog.askopenfilename()
    labelbt1 = Label(janela, text="{} CARREGADO".format(arquivogetnet), font="Arial 7")
    labelbt1.grid(column=0, row=3)


def ler2():
    global arquivocbb
    arquivocbb = filedialog.askopenfilename()
    labelbt2 = Label(janela, text="{} CARREGADO".format(arquivocbb), font="Arial 7")
    labelbt2.grid(column=0, row=4)

def truncate():
    connection = create_server_connection("localhost", "root", "wolf")
    usardb = "USE fluxodecaixa"
    truncate = "TRUNCATE TABLE getnet"
    truncate2 = "TRUNCATE TABLE bbrasil"
    truncate3 = "TRUNCATE TABLE results"
    execute_query(connection, usardb)
    execute_query(connection, truncate)
    execute_query(connection, truncate2)
    execute_query(connection, truncate3)

def start():
    a = Th(1)
    a.start()

class Th(Thread):
    def __init__(self, num):
        Thread.__init__(self)
        self.num = num


    def run(self):
        global arquivogetnet
        global pastadetrabalhogetnet
        global valoresgetnet
        global datastr
        global results


        pastadetrabalhogetnet = xlwings.Book(arquivogetnet)
        planilha = pastadetrabalhogetnet.sheets['Planilha1']

        pastadetrabalhocbb = xlwings.Book(arquivocbb)
        planilhacbb = pastadetrabalhocbb.sheets['Planilha1']

        last_row = planilha.range('B1').end('down').row
        last_row2 = planilha.range('A1').end('down').row

        last_row3 = planilhacbb.range('A1').end('down').row
        last_row4 = planilhacbb.range('E1').end('down').row

        for a in range(1, last_row3 + 1):
            valtemp = planilhacbb.range('A{}'.format(a)).value
            datascbb.append(valtemp)

        for b in range(1, last_row4 + 1):
            valtemp = planilhacbb.range('E{}'.format(b)).value
            valorescbb.append(valtemp)

        for i in range(1, last_row + 1):
            valtemp = planilha.range('B{}'.format(i)).value
            valoresgetnet.append(valtemp)

        for k in range(1, last_row2 + 1):
            valtemp = planilha.range('A{}'.format(k)).value
            datasgetnet.append(valtemp)

        for c in range(0, len(datasgetnet)):
            date = datetime.strptime(datasgetnet[c], '%d/%m/%Y').date()
            datastr.append(date)

        pastadetrabalhogetnet.close()
        pastadetrabalhocbb.close()
        connection = create_server_connection("localhost", "root", "wolf")
        print(datastr)

        for j in range(0, len(valoresgetnet)):
            usardb = "USE fluxodecaixa"
            inserir = "INSERT INTO getnet (dataatual, valor) VALUES ('{}', '{}')".format(datastr[j], valoresgetnet[j])
            execute_query(connection, usardb)
            execute_query(connection, inserir)

        for y in range(0, len(valorescbb)):
            usardb = "USE fluxodecaixa"
            inserir = "INSERT INTO bbrasil (dataatual, valor) VALUES ('{}', '{}')".format(datascbb[y], valorescbb[y])
            execute_query(connection, usardb)
            execute_query(connection, inserir)

        conn = mysql.connector.connect(
            host="localhost",
            user="root",
            password="wolf",
            database="fluxodecaixa"
        )


        usardbglobal = "USE fluxodecaixa"
        insertselect = "INSERT INTO results (dataatualbb, valorbb, somaacumuladabb) SELECT dataatual, valor, sum(valor) OVER (PARTITION BY dataatual ORDER BY dataatual) AS soma_acumulada FROM bbrasil GROUP BY dataatual, valor ORDER BY dataatual;"
        execute_query(connection, usardbglobal)
        execute_query(connection, insertselect)

        cursor = conn.cursor()
        cursor.execute("SELECT dataatualbb, somaacumuladabb FROM results;")
        results = cursor.fetchall()

        app = xlwings.App()
        workbook = app.books.add()
        sheet = workbook.sheets.active
        sheet.range('A1').value = results

        print(results)

        cursor.close()
        conn.close()




#INTERFACE
janela = Tk()
janela.title('FLUXO DE CAIXA')
janela.geometry("300x300")

Label1 = Label(janela, text='Insira as pastas de trabalho:', font="Arial 10 bold", justify=CENTER)
Label1.grid(column=0, row=0, padx=50, pady=10)

Botao1 = Button(janela, text='GETNET', font="Arial 10")
Botao1.grid(column=0, row=1, padx=10, pady=10)
Botao1.bind("<Button>", lambda e: ler1())

botao2 = Button(janela, text="COBRANCABB", font="Arial 10")
botao2.grid(column=0, row=2, padx=10, pady=10)
botao2.bind("<Button>", lambda e: ler2())

botao3 = Button(janela, text="GERAR CONTROLE", font="Arial 10")
botao3.grid(column=0, row=5, padx=10, pady=10)
botao3.bind("<Button>", lambda e: start())

truncatebtn = Button(janela, text="TRUNCATE", font="Arial 10")
truncatebtn.grid(column=0, row=6, padx=10, pady=10)
truncatebtn.bind("<Button>", lambda e: truncate())

janela.mainloop()


