import xlwings
from xlwings import *
import threading
import tkinter as tk
from tkinter import *
from tkinter import filedialog
from threading import Thread
import mysql.connector
from mysql.connector import Error
import datetime
from datetime import datetime, timedelta

arquivogetnet = None
pastadetrabalhogetnet = None


arquivocbb = None
pastadetrabalhocbb = None

arquivocontas = None

datagetnet = []
valorgetnet =[]

databb = []
valorbb = []

datacontas = []
valorcontas = []

datastr = []

valorcontasfloat = []

count = 2
count2 = 0

contdata = 3
contdata2 = 1

fimgetnet = []
fimbb = []
fimdespesas = []

colunadatas = []

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

def ler3():
    global arquivocontas
    arquivocontas = filedialog.askopenfilename()
    labelbt3 = Label(janela, text="{} CARREGADO".format(arquivocontas), font="Arial 7")
    labelbt3.grid(column=0, row=6)

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
        global count
        global count2
        global contdata
        global contdata2


        pastadetrabalhogetnet = xlwings.Book(arquivogetnet)
        planilha = pastadetrabalhogetnet.sheets['Planilha1']

        pastadetrabalhocbb = xlwings.Book(arquivocbb)
        planilhacbb = pastadetrabalhocbb.sheets['Planilha1']

        pastadetrabalhocontas = xlwings.Book(arquivocontas)
        planilhacontas = pastadetrabalhocontas.sheets[0]

        getnetdata = planilha.range('A1').end('down').row
        getnetvalor = planilha.range('B1').end('down').row

        bblastrow = planilhacbb.range('A1').end('down').row

        despesaslr = planilhacontas.range('A9').end('down').row

        for i in range(1, getnetdata + 1):
            data = planilha.range('A{}'.format(i)).value
            valor = planilha.range('B{}'.format(i)).value
            datagetnet.append(data)
            valorgetnet.append(valor)

        for i in range(1, bblastrow + 1):
            cell = planilhacbb.range('D{}'.format(i)).value
            if cell == None:
                data = planilhacbb.range('A{}'.format(i - 1)).value
                valor = planilhacbb.range('E{}'.format(i)).value
                databb.append(data)
                valorbb.append(valor)

        for i in range(9, despesaslr):
            cell = planilhacontas.range('B{}'.format(i)).value
            if cell == None:
                data = planilhacontas.range('A{}'.format(i - 1)).value
                valor = planilhacontas.range('F{}'.format(i)).value
                datacontas.append(data)
                valorcontas.append(valor)

        for i in range(0, len(datagetnet)):
            date = datetime.strptime(datagetnet[i], '%d/%m/%Y').date()
            datastr.append(date)

        for i in range(0, len(valorcontas)):
            valor = float(valorcontas[i])
            valorcontasfloat.append(valor)

        connection = create_server_connection("localhost", "root", "wolf")


        usardb = "USE fluxodecaixa;"
        execute_query(connection, usardb)

        for i in range(0, len(datastr)):
            inserir = "INSERT INTO getnet (data, valor) VALUES ('{}', '{}');".format(datastr[i], valorgetnet[i])
            execute_query(connection, inserir)

        for i in range(0, len(databb)):
            inserir = "INSERT INTO bbrasil (data, valor) VALUES ('{}', '{}')".format(databb[i], valorbb[i])
            execute_query(connection, inserir)

        for i in range(0, len(datacontas)):
            inserir = "INSERT INTO despesas (data, valor) VALUES ('{}', '{}')".format(datacontas[i], valorcontasfloat[i])
            execute_query(connection, inserir)

        x = datetime.now()

        app = xlwings.App()
        workbook = app.books.add()
        sheet = workbook.sheets.active
        sheet.range('A1').value = "DATA"
        sheet.range('B1').value = "CREDITOS BB"
        sheet.range('C1').value = "CREDITOS GETNET"
        sheet.range('D1').value = "DEBITOS"

        sheet.range('A2').value = x.strftime("%x")

        for i in range(0, 1095):
            x_incremento = x + timedelta(days=contdata2)
            sheet.range('A{}'.format(contdata)).value = x_incremento.strftime("%x")
            contdata += 1
            contdata2 += 1

        conn = mysql.connector.connect(
            host="localhost",
            user="root",
            password="wolf",
            database="fluxodecaixa"
        )

        cursor = conn.cursor()
        cursor.execute("SELECT * FROM getnet ORDER BY data")
        fimgetnet = cursor.fetchall()

        cursor2 = conn.cursor()
        cursor2.execute("SELECT * FROM bbrasil ORDER BY data")
        fimbb = cursor.fetchall()

        cursor3 = conn.cursor()
        cursor3.execute("SELECT * FROM despesas ORDER BY data")
        fimdespesas = cursor.fetchall()

        last_rowdatas = workbook.range('A2').end('down').row

        for i in range(0, last_rowdatas + 1):
            temp = workbook.range('A{}'.format(i)).value
            colunadatas.append(temp)

        print(colunadatas)

        pastadetrabalhogetnet.close()
        pastadetrabalhocbb.close()
        pastadetrabalhocontas.close()

#INTERFACE
janela = Tk()
janela.title('FLUXO DE CAIXA')
janela.geometry("500x300")

Label1 = Label(janela, text='Insira as pastas de trabalho:', font="Arial 10 bold", justify=CENTER)
Label1.grid(column=0, row=0, padx=150, pady=10)

Botao1 = Button(janela, text='GETNET', font="Arial 10")
Botao1.grid(column=0, row=1, padx=10, pady=10)
Botao1.bind("<Button>", lambda e: ler1())

botao2 = Button(janela, text="COBRANCABB", font="Arial 10")
botao2.grid(column=0, row=2, padx=10, pady=10)
botao2.bind("<Button>", lambda e: ler2())

botao4 = Button(janela, text="CONTAS A PAGAR", font="Arial 10")
botao4.grid(column=0, row=5, padx=10, pady=10)
botao4.bind("<Button>", lambda e: ler3())

botao3 = Button(janela, text="GERAR CONTROLE", font="Arial 10")
botao3.grid(column=0, row=7, padx=10, pady=10)
botao3.bind("<Button>", lambda e: start())

janela.mainloop()


