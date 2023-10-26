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

arquivosafra = None
pastadetrabalhosafra = None

arquivodeposito = None
pastadetrabalhodeposito = None

datadeposito = []
valordeposito = []

datagetnet = []
valorgetnet =[]

databb = []
valorbb = []

datacontas = []
valorcontas = []

datasafra = []
valorsafra = []

datastr = []

valorcontasfloat = []

datassafrareal = []

count = 2
count2 = 0
count3 = 2
count4 = 2
count5 = 2
count6 = 2
count7 = 2
count8 = 2

contdata = 3
contdata2 = 1

fimgetnet = []
fimbb = []
fimdespesas = []
fimsafrapay = []
fimdepositos = []

colunadatas = []

getnetdatafim = []

datadata = []

#FUNÇÃO PARA EXECUTAR COMANDOS NO MYSQL
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

#FUNÇÃO DE RETORNO DA EXECUÇÃO DAS QUERYS NO MYSQL
def execute_query(connection, query):
    cursor = connection.cursor()
    try:
        cursor.execute(query)
        connection.commit()
        print("Query successful")
    except Error as err:
        print(f"Error: '{err}'")

#LER O PRIMEIRO ARQUIVO
def ler1():
    global arquivogetnet
    arquivogetnet = filedialog.askopenfilename()
    labelbt1 = Label(janela, text="{} CARREGADO".format(arquivogetnet), font="Arial 7")
    labelbt1.grid(column=0, row=3)

#LER O SEGUNDO ARQUIVO
def ler2():
    global arquivocbb
    arquivocbb = filedialog.askopenfilename()
    labelbt2 = Label(janela, text="{} CARREGADO".format(arquivocbb), font="Arial 7")
    labelbt2.grid(column=0, row=4)

#LER O TERCEIRO ARQUIVO
def ler3():
    global arquivocontas
    arquivocontas = filedialog.askopenfilename()
    labelbt3 = Label(janela, text="{} CARREGADO".format(arquivocontas), font="Arial 7")
    labelbt3.grid(column=0, row=6)

#LER O QUARTO ARQUIVO
def ler4():
    global arquivosafra
    arquivosafra = filedialog.askopenfilename()
    labelbt4 = Label(janela, text="{} CARREGADO".format(arquivosafra), font="Arial 7")
    labelbt4.grid(column=0, row=8)

#LER O QUINTO ARQUIVO
def ler5():
    global arquivodeposito
    arquivodeposito = filedialog.askopenfilename()
    labelbt5 = Label(janela, text="{} CARREGADO".format(arquivodeposito), font="Arial 7")
    labelbt5.grid(column=0, row=10)

#THREADING
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
        global count3
        global count4
        global count5
        global datadata
        global count6
        global count7
        global count8

        #VARIÁVEIS DE ARMAZENAMENTO PARA ARQUIVOS INSERIDOS E IDENTIFICAÇÃO DE PLANILHAS
        pastadetrabalhogetnet = xlwings.Book(arquivogetnet)
        planilha = pastadetrabalhogetnet.sheets['Planilha1']

        pastadetrabalhocbb = xlwings.Book(arquivocbb)
        planilhacbb = pastadetrabalhocbb.sheets['Planilha1']

        pastadetrabalhocontas = xlwings.Book(arquivocontas)
        planilhacontas = pastadetrabalhocontas.sheets[0]

        pastadetrabalhosafra = xlwings.Book(arquivosafra)
        planilhasafra = pastadetrabalhosafra.sheets[0]

        pastadetrabalhodeposito = xlwings.Book(arquivodeposito)
        planilhadeposito = pastadetrabalhodeposito.sheets[0]

        #REFERÊNCIA DE LOOP PARA A LEITURA DOS DADOS
        getnetdata = planilha.range('A1').end('down').row

        bblastrow = planilhacbb.range('A1').end('down').row

        despesaslr = planilhacontas.range('A1').end('down').row

        safralr = planilhasafra.range('F1').end('down').row

        depositolr = planilhadeposito.range('H1').end('down').row

        #LENDO DADOS EM GETNET
        for i in range(1, getnetdata + 1):
            data = planilha.range('A{}'.format(i)).value
            valor = planilha.range('B{}'.format(i)).value
            datagetnet.append(data)
            valorgetnet.append(valor)

        #LENDO DADOS EM BANCO DO BRASIL
        for i in range(1, bblastrow + 1):
            cell = planilhacbb.range('D{}'.format(i)).value
            if cell == None:
                data = planilhacbb.range('A{}'.format(i - 1)).value
                valor = planilhacbb.range('E{}'.format(i)).value
                databb.append(data)
                valorbb.append(valor)

        #LENDO DADOS EM DESPESAS
        for i in range(1, despesaslr):
            cell = planilhacontas.range('B{}'.format(i)).value
            if cell == None:
                data = planilhacontas.range('A{}'.format(i - 1)).value
                valor = planilhacontas.range('F{}'.format(i)).value
                datacontas.append(data)
                valorcontas.append(valor)

        # LENDO DADOS EM SAFRAPAY
        for i in range(1, safralr):
            cell = planilhasafra.range('E{}'.format(i)).value
            if cell == None:
                data = planilhasafra.range('F{}'.format(i - 1)).value
                valor = planilhasafra.range('L{}'.format(i)).value
                datasafra.append(data)
                valorsafra.append(valor)

        # LENDO DADOS EM DEPOSITOS
        for i in range(1, depositolr):
            cell = planilhadeposito.range('G{}'.format(i)).value
            if cell == None:
                data = planilhadeposito.range('H{}'.format(i - 1)).value
                valor = planilhadeposito.range('I{}'.format(i)).value
                datadeposito.append(data)
                valordeposito.append(valor)

        #TRATAMENTO DE DADOS
        for i in range(0, len(datagetnet)):
            date = datetime.strptime(datagetnet[i], '%d/%m/%Y').date()
            datastr.append(date)

        for i in range(0, len(datasafra)):
            temp = datasafra[i].replace('.', '/')
            date = datetime.strptime(temp, '%d/%m/%Y').date()
            datassafrareal.append(date)

        for i in range(0, len(valorcontas)):
            valor = float(valorcontas[i])
            valorcontasfloat.append(valor)

        #CONEXÃO PADRÃO COM O SERVIDOR MYSQL
        connection = create_server_connection("localhost", "root", "wolf")


        usardb = "USE fluxodecaixa;"
        execute_query(connection, usardb)

        #INSERÇÃO DE DADOS NO SERVIDOR MYSQL
        for i in range(0, len(datastr)):
            inserir = "INSERT INTO getnet (data, valor) VALUES ('{}', '{}');".format(datastr[i], valorgetnet[i])
            execute_query(connection, inserir)

        for i in range(0, len(databb)):
            inserir = "INSERT INTO bbrasil (data, valor) VALUES ('{}', '{}')".format(databb[i], valorbb[i])
            execute_query(connection, inserir)

        for i in range(0, len(datacontas)):
            inserir = "INSERT INTO despesas (data, valor) VALUES ('{}', '{}')".format(datacontas[i], valorcontasfloat[i])
            execute_query(connection, inserir)

        for i in range(0, len(datasafra)):
            inserir = "INSERT INTO safra (data, valor) VALUES ('{}', '{}')".format(datassafrareal[i], valorsafra[i])
            execute_query(connection, inserir)

        for i in range(0, len(datadeposito)):
            inserir = "INSERT INTO depositos (data, valor) VALUES ('{}', '{}')".format(datadeposito[i], valordeposito[i])
            execute_query(connection, inserir)

        #DATA ATUAL
        x = datetime.now()

        #FORMATAÇÃO DA PLANILHA RESULTADO
        app = xlwings.App()
        workbook = app.books.add()
        sheet = workbook.sheets.active
        sheet.range('A1').value = "DATA"
        sheet.range('B1').value = "CREDITOS BB"
        sheet.range('C1').value = "CREDITOS GETNET"
        sheet.range('D1').value = "CREDITOS SAFRAPAY"
        sheet.range('E1').value = "DEPOSITOS"
        sheet.range('F1').value = "DEBITOS"


        sheet.range('A2').value = x.strftime("%x")

        #INSERINDO DATAS SEQUENCIAIS NA COLUNA A NO PERÍODO DE 4 ANOS
        for i in range(0, 1460):
            x_incremento = x + timedelta(days=contdata2)
            sheet.range('A{}'.format(contdata)).value = x_incremento.strftime("%x")
            contdata += 1
            contdata2 += 1

        #CONEXÃO COM O SERVIDOR MYSQL
        conn = mysql.connector.connect(
            host="localhost",
            user="root",
            password="wolf",
            database="fluxodecaixa"
        )

        #SELECIONANDO OS DADOS DO BANCO DE DADOS E ARMAZENANDO EM VETORES
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM getnet ORDER BY data")
        fimgetnet = cursor.fetchall()

        cursor2 = conn.cursor()
        cursor2.execute("SELECT * FROM bbrasil ORDER BY data")
        fimbb = cursor.fetchall()

        cursor3 = conn.cursor()
        cursor3.execute("SELECT * FROM despesas ORDER BY data")
        fimdespesas = cursor.fetchall()

        cursor8 = conn.cursor()
        cursor8.execute("SELECT * FROM safra ORDER BY data")
        fimsafrapay = cursor.fetchall()

        cursor9 = conn.cursor()
        cursor9.execute("SELECT * FROM depositos ORDER BY data")
        fimdepositos = cursor.fetchall()

        #IDENTIFICANDO A ULTIMA LINHA COM DADOS NA COLUNA DE DATAS
        last_rowdatas = sheet.range('A2').end('down').row


        #ARMAZENANDO COLUNA DE DATAS
        for i in range(2, last_rowdatas + 1):
            temp = sheet.range('A{}'.format(i)).value
            colunadatas.append(temp)

        #TRATANDO OS DADOS DE DATAS PARA O TIPO DATE
        for i in range(0, len(colunadatas)):
            temp = colunadatas[i]
            data = temp.date()
            datadata.append(data)

        #GETNET
        local = []
        for i in range(0, len(fimgetnet)):
            temp = fimgetnet[i][1]
            local.append(temp)

        for i in range(0, len(datadata)):
            valor = datadata[i]
            if valor in local:
                indice = local.index(valor)
                sheet.range('C{}'.format(count3)).value = fimgetnet[indice][2]
                count3 += 1
            else:
                print('{} Not in list'.format(valor))
                count3 += 1


        #BANCO DO BRASIL
        local2 = []
        for i in range(0, len(fimbb)):
            temp = fimbb[i][1]
            local2.append(temp)

        for i in range(0, len(datadata)):
            valor = datadata[i]
            if valor in local2:
                indice = local2.index(valor)
                sheet.range('B{}'.format(count5)).value = fimbb[indice][2]
                count5 += 1
            else:
                print('{} Not in list'.format(valor))
                count5 += 1

        #DESPESAS
        local3 = []
        for i in range(0, len(fimdespesas)):
            temp = fimdespesas[i][1]
            local3.append(temp)

        for i in range(0, len(datadata)):
            valor = datadata[i]
            if valor in local3:
                indice = local3.index(valor)
                sheet.range('F{}'.format(count6)).value = fimdespesas[indice][2]
                count6 += 1
            else:
                print('{} Not in list'.format(valor))
                count6 += 1

        #SAFRAPAY
        local4 = []
        for i in range(0, len(fimsafrapay)):
            temp = fimsafrapay[i][1]
            local4.append(temp)

        for i in range(0, len(datadata)):
            valor = datadata[i]
            if valor in local4:
                indice = local4.index(valor)
                sheet.range('D{}'.format(count7)).value = fimsafrapay[indice][2]
                count7 += 1
            else:
                print('{} Not in list'.format(valor))
                count7 += 1

        #DEPOSITOS
        local5 = []
        for i in range(0, len(fimdepositos)):
            temp = fimdepositos[i][1]
            local5.append(temp)

        for i in range(0, len(datadata)):
            valor = datadata[i]
            if valor in local5:
                indice = local5.index(valor)
                sheet.range('E{}'.format(count8)).value = fimsafrapay[indice][2]
                count8 += 1
            else:
                print('{} Not in list'.format(valor))
                count8 += 1

        cursor4 = conn.cursor()
        cursor4.execute("TRUNCATE getnet;")
        cursor5 = conn.cursor()
        cursor5.execute("TRUNCATE bbrasil;")
        cursor6 = conn.cursor()
        cursor6.execute("TRUNCATE despesas;")
        cursor7 = conn.cursor()
        cursor7.execute("TRUNCATE safra;")
        cursor10 = conn.cursor()
        cursor10.execute("TRUNCATE depositos;")

        pastadetrabalhogetnet.close()
        pastadetrabalhocbb.close()
        pastadetrabalhocontas.close()
        pastadetrabalhodeposito.close()


#INTERFACE
janela = Tk()
janela.title('FLUXO DE CAIXA')
janela.geometry("500x500")

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
botao3.grid(column=0, row=11, padx=10, pady=10)
botao3.bind("<Button>", lambda e: start())

botao5 = Button(janela, text="SAFRAPAY", font="Arial 10")
botao5.grid(column=0, row=7, padx=10, pady=10)
botao5.bind("<Button>", lambda e: ler4())

botao6 = Button(janela, text="DEPOSITOS", font="Arial 10")
botao6.grid(column=0, row=9, padx=10, pady=10)
botao6.bind("<Button>", lambda e: ler5())

janela.mainloop()


