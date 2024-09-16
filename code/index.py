from tkinter import *
from tkinter import messagebox
from tkinter.filedialog import askopenfilenames, askopenfilename, asksaveasfilename
import tabula as tb
import pandas as pd
from unidecode import unidecode
import string
import os

import copy

from openpyxl import *
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment
import sys
from datetime import *
from xlsxwriter import *
import locale

locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')

window = Tk()

class Arquivo:
    def __init__(self):
        self.tipos_validos = ''

    def validar_entrada(self, caminho):
        if any(c not in string.ascii_letters for c in caminho):
            caminho = self.formato_ascii(caminho)

        self.__tipo(caminho)
        return caminho

    def __tipo(self, caminho):
        if caminho[len(caminho) -3 :] != self.tipos_validos:
            ultima_barra = caminho.rfind('/')
            raise Exception(
                f'Formato inválido do arquivo: {caminho[ultima_barra+1:]}')

    def formato_ascii(self, caminho):
        caminho_uni = unidecode(caminho)
        os.renames(caminho, caminho_uni)
        return caminho_uni
    
    def envio_invalido(self):
        return True if len(self.caminho) == 0 else False

class Matriz(Arquivo):
    def __init__(self):
        super().__init__()
        self.tipos_validos = 'lsx'
        self.caminho = ''

    def inserir(self, label):
        try:
            caminho = askopenfilename()

            if caminho == '':
                return None

            caminho_validado = self.validar_entrada(caminho)
            label['text'] = caminho[caminho.rfind('/') +1:]

            self.caminho = caminho_validado

        except PermissionError:
            messagebox.showerror(title='Aviso', message= 'O arquivo selecionado apresenta-se aberto em outra janela, favor fecha-la')
        except FileExistsError:
            messagebox.showerror(title='Aviso', message= 'O arquivo selecionado já apresenta uma versão sem acento, favor usar tal versão ou apagar uma delas')
        except Exception as error:
            messagebox.showerror(title='Aviso', message= error)

    def ler(self):
        return pd.read_excel(self.caminho, na_filter=False, usecols='A:B')
    
    def load(self):
        return load_workbook(self.caminho)
    
class Recibo(Arquivo):
    def __init__(self):
        super().__init__()
        self.tipos_validos = 'pdf'
        self.caminho = []

    def get_caminho(self):
        return self.caminho

    def inserir(self, label):
        try:
            caminhos = askopenfilenames()

            if caminhos == '':
                return None

            label.delete(0, END)
            caminhos_validados = []
            for item in caminhos:
                caminhos_validados.append(self.validar_entrada(item))
                label.insert(END, f'{item[item.rfind('/') +1:]}\n')

            self.caminho = caminhos_validados.copy()

        except PermissionError:
            messagebox.showerror(title='Aviso', message= 'O arquivo selecionado apresenta-se em aberto em outra janela, favor fecha-la')
        except FileExistsError:
            messagebox.showerror(title='Aviso', message= 'O arquivo selecionado já apresenta uma versão sem acento, favor usar tal versão ou apagar uma delas')
        except Exception as error:
            messagebox.showerror(title='Aviso', message= error)

class Writer:
    def __init__(self, df, titulo):
        self.df = df
        self.titulo = titulo

        self.LIN_INDEX = 6
        self.lin_data = self.LIN_INDEX + 1
        self.dif_data = 2
        self.dif_cnpj = 1

    def abrir(self):
        messagebox.showinfo(title='Aviso', message='Abrindo o arquivo gerado!')

        os.startfile(self.file+'.xlsx')
    
    def colunas(self):
        for col_index, columns in enumerate(list(self.df.columns)):
            self.ws.write(self.LIN_INDEX, col_index, columns,self.wb.add_format({'bold':True,'top':2, 'bg_color':'#a7b8ab','underline':True, 'align':'center'}))

    def preencher_matriz(self):
        for index, row in self.df_matriz.iterrows():
            #Adciona Nome e CNPJ apenas
            for col_index, valor in enumerate(row):
                self.ws.write(index + self.lin_data, col_index, valor,\
                    self.wb.add_format({'border':3, 'align':'center'}))
            
            #Termina o df com espaços vazios
            self.espacos_vazios(index)

    def espacos_vazios(self, index):
        for col_index in range(len(self.df.columns) - self.dif_data):
            self.ws.write(index + self.lin_data, col_index + self.dif_data, '',\
                self.wb.add_format({'border':3}))

    def preencher_data(self, excluido, repetido):
        encontrados = []
        for index_recibo, row_recibo in self.df.iterrows():
            achado = False
            #print(f'{row_recibo['CNPJ']} - CNPJ procurado')
            for index_matriz, row_matriz in self.df_matriz.iterrows():
                #print(f'{row_matriz['CNPJ']} - opções')
                if row_recibo['CNPJ'] == row_matriz['CNPJ']:
                    achado = True
                    if row_recibo['CNPJ'] in encontrados:
                        repetido.add_data(row_recibo)
                    else:
                        encontrados.append(row_recibo['CNPJ'])
                        for col_index, valor in enumerate(row_recibo):
                            self.ws.write(index_matriz + self.lin_data, col_index, valor, self.wb.add_format({'border':3, 'align':'center'}))

                    break

            if achado == False:
                excluido.add_data(row_recibo)
                    
    def nomear_arq(self):
        nome_arq = asksaveasfilename(title='Favor selecionar a pasta onde será salvo', filetypes=((".xlsx","*.xlsx"),))

        if nome_arq == '':
            if messagebox.askyesno(title='Aviso', message= 'Deseja cancelar esta operação?') == True:
                raise Exception ('Operação cancelada!')
            else:
                return self.nomear_arq()
            
        return nome_arq    
    
class Relatorio(Writer):
    def __init__(self, df, df_matriz, titulo):
        super().__init__(df, titulo)
        self.df_matriz = df_matriz

        self.file = self.nomear_arq()

        self.wb = Workbook(self.file + '.xlsx')
        self.ws = self.wb.add_worksheet('Relacionados')
        self.ws.set_column('A:A', 40)
        self.ws.set_column('B:B', 20)
        self.ws.set_column('C:C', 15)
        self.ws.set_column('D:D', 20)
        self.ws.set_column('E:E', 15)
        self.ws.set_column('F:F', 20)
        self.ws.set_column('G:G', 20)

        self.excluidos = Adcional(self.df, self.titulo, self.wb)
        self.repetidos = Adcional(self.df, self.titulo, self.wb)


    def gerar_arquivo(self):
        self.__cabecalho()
        self.__table_ref()
        self.colunas()
        self.preencher_matriz()
        
        self.preencher_data(self.excluidos, self.repetidos)

        if self.excluidos.qnt_itens() != 0:
            messagebox.showinfo(title='Aviso', message= f'{self.excluidos.qnt_itens()} empresas foram inseridas na aba "Não relacionadas" por não constarem na matriz')

            ws = self.wb.add_worksheet('Não Relacionados')
            ws.set_column('A:A', 40)
            ws.set_column('B:B', 20)
            ws.set_column('C:C', 15)
            ws.set_column('D:D', 20)
            ws.set_column('E:E', 15)
            ws.set_column('F:F', 20)
            ws.set_column('G:G', 20)
            self.excluidos.preencher(ws,'yellow')

        if self.repetidos.qnt_itens() != 0:
            messagebox.showinfo(title='Aviso', message= f'{self.repetidos.qnt_itens()} empresas foram inseridos em duplicidade. A segunda cópia foi inserida na aba "Repetidos"')

            ws = self.wb.add_worksheet('Repetidos')
            ws.set_column('A:A', 40)
            ws.set_column('B:B', 20)
            ws.set_column('C:C', 15)
            ws.set_column('D:D', 20)
            ws.set_column('E:E', 15)
            ws.set_column('F:F', 20)
            ws.set_column('G:G', 20)
            self.repetidos.preencher(ws ,'cyan')

        self.wb.close()    

    def __cabecalho(self):
        self.ws.write(0,0,f'RELATÓRIO DE CONFERÊNCIA {self.titulo}',\
            self.wb.add_format({'bold': True, 'font_size': 26}))

        self.ws.write(2,0,'Competência',\
            self.wb.add_format({'bold':True,'align':'right','font_size': 16}))
        self.ws.write(2,1, self.data_confe())

        self.ws.write(3,0,'Data Entrega',\
            self.wb.add_format({'bold':True,'align':'right','font_size': 16}))
        self.ws.write(3,1, datetime.now().strftime("%d/%m/%Y"))

    def data_confe(self):
        data = f'{datetime.now().month - 1}/{datetime.now().year}'
        data_format = datetime.strptime(data, '%m/%Y')
        return data_format.strftime("%B/%Y".capitalize())

    def __table_ref(self):
        tam_df = len(self.df_matriz)+7
        ref = {
            'Obrigadas:': f'=COUNTA($A8:$A{tam_df})',
            'Entregues:': f'=COUNTA($C8:$C{tam_df})',
            'Não Entregues:': '=$E3 - $E4'
        }

        #f'=COUNTIF($D8:$D{tam_df}, "ENVIADO")'
        for index, text in enumerate(ref.items()):
            self.ws.write(index+2,3, text[0],\
                self.wb.add_format({'bold':True,'border':1,'align':'right'}))
            
            self.ws.write(index+2,4, text[1],\
                self.wb.add_format({'border':1,'align':'center'}))

class Incremento(Writer):
    def __init__(self, df, df_relatorio, wb_completo, titulo):
        super().__init__(df, titulo)

        titulo_arq = df_relatorio.columns[0]
        if self.titulo not in titulo_arq:
            raise Exception('O relatório inserido é de uma obrigação diferente do recibo em questão')

        self.df_relatorio = df_relatorio
        self.df_relatorio = self.df_relatorio.drop([0,1,2,3,4,5,6])\
            .reset_index(drop=True)
        
        self.wb = wb_completo
        self.ws = self.wb['Relacionados']
        self.excluidos = Adcional(self.df, self.titulo, self.wb)
        self.repetidos = Adcional(self.df, self.titulo, self.wb)

    def gerar_arquivo(self):
        self.data()

        if self.excluidos.qnt_itens() != 0:
            messagebox.showinfo(title='Aviso', message= f'{self.excluidos.qnt_itens()} empresas foram inseridas na aba "Não relacionadas" por não constarem na matriz')

            self.gerar_aba('Não Relacionados', 'yellow')
        
        if self.repetidos.qnt_itens() != 0:
            messagebox.showinfo(title='Aviso', message= f'{self.repetidos.qnt_itens()} empresas foram inseridos em duplicidade. A segunda cópia foi inserida na aba "Repetidos"')

            self.gerar_aba('Repetidos', 'cyan')

        self.file = self.nomear_arq()

        self.wb.save(self.file+'.xlsx')

    def data(self):
        for index_recibo, row_recibo in self.df.iterrows():
            achado = False
            #print(f'{row_recibo['CNPJ']} - CNPJ procurado')
            for index_matriz, row_matriz in self.df_relatorio.iterrows():
                #print(f'{row_matriz['CNPJ']} - opções')
                if row_recibo.iloc[1] == row_matriz.iloc[1]:
                    achado = True
                    if row_recibo.iloc[2] != '':
                        self.repetidos.add_data(row_recibo)
                    for col_index, valor in enumerate(row_recibo):
                        self.ws.cell\
                            (index_matriz + self.lin_data+2, col_index+1, valor).alignment = Alignment(horizontal='center')
                    break

            if achado == False:
                self.excluidos.add_data(row_recibo)
    
    def gerar_aba(self, titulo, cor):
        ws = self.wb.create_sheet(titulo)
        for index, valor in enumerate([40,20,15,20,15,20,20],1):
            ws.column_dimensions[get_column_letter(index)].width = valor
        self.repetidos.preencher(ws,cor)

class Adcional(Writer):
    def __init__(self, df, titulo, wb):
        super().__init__(df, titulo)
        self.wb = wb
        self.data = []

    def preencher(self, ws, cor):
        self.ws = ws
        self.__titulo()
        self.colunas()
        self.__excluidos(cor)

    def __titulo(self):
        self.ws.write(0,0,f'EMPRESAS NÃO RELACIONADAS {self.titulo}', self.wb.add_format({'bold': True, 'font_size': 26}))

    def __excluidos(self, cor):
        for index_recibo, row_recibo in enumerate(self.data):
            for col_index, valor in enumerate(row_recibo):
                self.ws.write(index_recibo + self.lin_data, col_index, valor, self.wb.add_format({'border':3, 'align':'center', 'bg_color':cor}))

    def add_data(self, row):
        self.data.append(row)

    def qnt_itens(self):
        return len(self.data)

class Competencia:
    def __init__(self):
        self.nome_emp = []
        self.cnpj = []
        self.referencia = []
        self.data = []
        self.hora = []

    def to_string(self):
        return self.titulo

class Des(Competencia):
    def __init__(self):
        super().__init__()
        self.servicos = []
        self.titulo = 'DES'

    def add_linha(self, arquivo):
        tabela = tb.read_pdf(arquivo, pages= 1, stream= True,\
                        relative_area=True, area= [10,0,59,68])[0]
        
        tabela.fillna('', inplace=True)

        ##Nome Emp
        self.nome_emp.append(tabela.iloc[1,0].replace('Nome/Razão Social: ',''))

        ##CNPJ
        if tabela.iloc[0,1] == '':
            self.cnpj.append(tabela.iloc[0,0][45:])
        else:
            self.cnpj.append(tabela.iloc[0,1])

        ##Ref.
        self.referencia.append(tabela.iloc[3,0]\
            .replace('Referência: ','')\
                    .replace(' No Protocolo:',''))

        ##Data e Hora
        col_dthr = tabela.iloc[4,0].replace('Data/Hora de Entrega: ','')\
                        .replace(' Regime de Tributação:','')

        self.data.append(col_dthr[:10])

        self.hora.append(col_dthr[10:])

        ##Serviços declarados.
        self.servicos.append(tabela.iloc[15,0]\
            .replace('Total de Serviços Declarados: ','')\
                .replace('Base de Cálculo S/ Ret',''))

    def gerar_df(self):
        return pd.DataFrame({
            'Nome Empresa': self.nome_emp,
            'CNPJ': self.cnpj, 
            'Referência': self.referencia, 
            'Data Entrega': self.data,
            'Hora Entrega': self.hora,
            'Serviços Tomados': self.servicos
            })

class Reinf(Competencia):
    def __init__(self):
        super().__init__()
        self.num_dom = []
        self.situacao = []
        self.titulo = 'EFD REINF'

    def add_linha(self, arquivo):
        arquivo = tb.read_pdf(arquivo, pages=1, stream=True,\
                        relative_area=True ,area=[5,0,100,100])[0]

        tabela = arquivo.loc[arquivo['Unnamed: 2'] == 'R-2099 - Fechamento dos Eventos Periódicos']

        prim_linha = tabela.iloc[0,0]

        ##Num Domínio
        self.num_dom.append(prim_linha[:prim_linha.find('-')-1])

        ##Nome empresa
        self.nome_emp.append(prim_linha[prim_linha.find('-')+2:])

        ##CNPJ
        self.cnpj.append(tabela.iloc[0,1][:18])

        ##Ref
        self.referencia.append(tabela.iloc[0,2])

        ##Situação
        self.situacao.append(tabela.iloc[0,5].replace('Sucesso','ENVIADO'))

        ##Data e Hora
        col_dthr = tabela.iloc[0,7][:18]

        self.data.append(col_dthr[:10])

        self.hora.append(col_dthr[12:])

    def gerar_df(self):
        return pd.DataFrame({
            'Nome Empresa': self.nome_emp,
            'CNPJ': self.cnpj, 
            'Num. Domínio': self.num_dom,
            'Referência': self.referencia,
            'Situação': self.situacao, 
            'Data Entrega': self.data,
            'Hora Entrega': self.hora,
            })
    
class Contribuicoes(Competencia):
    def __init__(self):
        super().__init__()
        self.titulo = 'EFD CONTRIBUIÇÕES'

    def add_linha(self, arquivo):
        tabela = tb.read_pdf(arquivo, pages=1, stream=True,\
                        relative_area=True ,area=[10,0,100,100])[0]

        dif = 0
        dif_dthr = 0

        if 'IDENTIFICAÇÃO' not in tabela.iloc[0,0]:
            dif = 1
            dif_dthr = 20

        ##Nome Empresa
        self.nome_emp.append(tabela.iloc[1 + dif,0].replace('Contribuinte: ',''))

        ##CNPJ
        self.cnpj.append(tabela.iloc[2 + dif,0][6:24])

        ##Ref
        self.referencia.append(tabela.iloc[4 + dif,0][37:])

        ##Data e Hora
        col_dthr = tabela.iloc[28,0]

        self.data.append(col_dthr[3 + dif_dthr :13 + dif_dthr])

        self.hora.append(col_dthr[17 + dif_dthr:])

    def gerar_df(self):
        return pd.DataFrame({
            'Nome Empresa': self.nome_emp,
            'CNPJ': self.cnpj, 
            'Referência': self.referencia,
            'Data Entrega': self.data,
            'Hora Entrega': self.hora,
            })
    
class SimplesNacional(Competencia):
    def __init__(self):
        super().__init__()
        self.valor = []
        self.anexo = []
        self.titulo = 'SIMPLES NACIONAL'

    def add_linha(self, arquivo):
        tabela = tb.read_pdf(arquivo, pages=1, stream=True,\
                        relative_area=True ,area=[0.5,0,100,100])[0]

        ##Nome Empresa
        self.nome_emp.append(tabela.iloc[18,0].replace('Estabelecimento: ','')[3:])

        ##CNPJ
        self.cnpj.append(tabela.iloc[0,0].replace('CNPJ: ',''))

        ##Ref
        self.referencia.append(tabela.iloc[4,0].replace('Período: ',''))

        ##Data
        self.data.append(str(tabela.iloc[0,1]).replace('Emissão: ',''))

        ##Valor Tributo
        self.valor.append(tabela.iloc[len(tabela)-2,0].replace('Simples Nacional a recolher: ',''))

        ##Anexo
        self.anexo.append(tabela.iloc[19,0][7:15].strip())

    def gerar_df(self):
        return pd.DataFrame({
            'Nome Empresa': self.nome_emp,
            'CNPJ': self.cnpj, 
            'Referência': self.referencia,
            'Data Entrega': self.data,
            'Valor Tributo': self.valor,
            'Anexo': self.anexo
            })

class DCTF(Competencia):
    def __init__(self):
        super().__init__()
        self.valor = []
        self.titulo = 'DCTF'

    def sum_saldos(self, arquivo):
        valor_valido = []
        for row in arquivo.iloc[10:21,0]:
            pos_vir = row.find(',')
            valor = row[pos_vir+3:].replace('.','').replace(',','.').strip()
            if valor[:4] != '0.00':
                valor_valido.append(float(valor))
        if len(valor_valido) != 0:
            return sum(valor_valido)
        return 0

    def add_linha(self, arquivo):
        tabela = tb.read_pdf(arquivo, pages=1, stream=True, guess=False)[0]

        ##Nome
        self.nome_emp.append(tabela.iloc[4,0].replace('Nome Empresarial: ',''))

        ##CNPJ
        self.cnpj.append(tabela.iloc[3,0][6:24])

        ##Ref
        self.referencia.append(tabela.iloc[3,0][34:])

        data_row = tabela.iloc[54,0]\
            .replace('exigido este número de recibo: em ','')

        ##Data
        self.data.append(data_row[:10])

        ##Hora
        self.hora.append(data_row[13:])

        ##Valor Tributo
        self.valor.append(f'{self.sum_saldos(tabela):.2f}')

    def gerar_df(self):
        return pd.DataFrame({
            'Nome Empresa': self.nome_emp,
            'CNPJ': self.cnpj,
            'Referência': self.referencia,
            'Data': self.data,
            'Hora': self.hora,
            'Saldo em Aberto': self.valor
            })

class App:
    def __init__(self):
        self.window = window
        self.recibos = Recibo()
        self.matriz = Matriz()

        self.ref = {
            'des' : Des(),
            'reinf' : Reinf(),
            'contribuicoes': Contribuicoes(),
            'contribuições' : Contribuicoes(),
            'simples nacional': SimplesNacional(),
            'sn': SimplesNacional(),
            'dctf': DCTF()
        }

        self.tela()
        self.index()
        window.mainloop()

    def tela(self):
        self.window.configure(background='darkblue')
        self.window.resizable(False,False)
        self.window.geometry('860x500')
        self.window.iconbitmap(self.resource_path('imgs\\conf-icon.ico'))
        self.window.title('Conversor de Extrato')

    def resource_path(self,relative_path):
        base_path = getattr(
            sys,
            '_MEIPASS',
            os.path.dirname(os.path.abspath(__file__)))
        return os.path.join(base_path, relative_path)

    def index(self):
        self.index = Frame(self.window, bd=4, bg='lightblue')
        self.index.place(relx=0.05,rely=0.05,relwidth=0.9,relheight=0.9)

        #Titulo
        Label(self.index, text='Conferência Automática', background='lightblue', font=('arial',30,'bold')).place(relx=0.23,rely=0.18,relheight=0.15)

        #Logo
        self.logo = PhotoImage(file=self.resource_path('imgs\\deltaprice-hori.png'))
        
        self.logo = self.logo.subsample(4,4)
        
        Label(self.window, image=self.logo, background='lightblue', border=0)\
            .place(relx=0.205,rely=0.05,relwidth=0.7,relheight=0.2)

        self.valIncrement = BooleanVar()

        self.valIncrement.set(False)

        Radiobutton(self.index, text="Criar novo Relatório", value=False, variable= self.valIncrement).place(relx=0.45,rely=0.33)

        Radiobutton(self.index, text="Incrementar em Relatório antigo", value=True, variable= self.valIncrement).place(relx=0.65,rely=0.33)

        #Labels e Entrys
        ###########Matriz
        Label(self.index, text='Insira aqui a Matriz/Referência:',\
            background='lightblue', font=(10))\
                .place(relx=0.15,rely=0.33)

        self.matLabel = Label(self.index)
        self.matLabel.config(font=("Arial", 12, 'bold italic'), anchor= 's')
        self.matLabel.place(relx=0.21,rely=0.4,relwidth=0.7, relheight=0.055)
        
        Button(self.index, text='Enviar',\
            command = lambda: self.matriz.inserir(self.matLabel))\
                .place(relx=0.15,rely=0.4,relwidth=0.06,relheight=0.055)

        ###########Arquivo
        Label(self.index, text='Insira aqui os Recibos:',\
            background='lightblue', font=(10))\
                .place(relx=0.15,rely=0.48)

        self.arqLabel = Listbox(self.index, border= 0)
        self.arqLabel.config(font=("Arial", 8, 'bold italic'))
        self.arqLabel.place(relx=0.21,rely=0.55,relwidth=0.675, relheight=0.2)

        self.barra = Scrollbar(self.index, command= self.arqLabel.yview)\
            .place(relx=0.875,rely=0.55,relwidth=0.03, relheight=0.2)
        
        self.arqLabel.config(yscrollcommand= self.barra)
        
        Button(self.index, text='Enviar',\
            command = lambda: self.recibos.inserir(self.arqLabel))\
                .place(relx=0.15,rely=0.55,relwidth=0.06,relheight=0.055)

        ###########EFD
        Label(self.index, text='Caso o nome da obrigação assesória não constar no nome do arquivo',\
            background='lightblue', font=("Arial", 12, 'bold italic'))\
                .place(relx=0.15,rely=0.775)

        Label(self.index, text='Escolha a obrigação:',\
            background='lightblue', font=(10))\
                .place(relx=0.15,rely=0.825)
        
        self.declaracaoEntry = StringVar()

        self.declaracaoEntryOpt = ["DES", "REINF", "CONTRIBUIÇÕES", "SIMPLES NACIONAL", "DCTF"]

        self.declaracaoEntry.set('Escolha aqui')

        self.popup = OptionMenu(self.index, self.declaracaoEntry, *self.declaracaoEntryOpt)\
            .place(relx=0.375,rely=0.835,relwidth=0.2,relheight=0.06)

        #Botão enviar
        Button(self.index, text='Gerar Conferencia',\
            command= lambda: self.executar())\
                .place(relx=0.65,rely=0.85,relwidth=0.25,relheight=0.12)
        
    def declaracao(self):
        itens_label = self.arqLabel.get(0,END)
        lista_declara = []

        primeiro_obj = copy.deepcopy(self.declaracao_valid(itens_label[0]))

        for itens in itens_label:
            lista_declara.append(type(self.declaracao_valid(itens)))

        if len(set(lista_declara)) != 1:
            raise Exception('Nem todos elementos são da mesma obrigação, favor selecionar tipo')

        return primeiro_obj
    
    def declaracao_valid(self, valor):
        if self.declaracaoEntry.get() != 'Escolha aqui':
            for key, obj in self.ref.items():
                if key in self.declaracaoEntry.get().lower():
                    return obj
        else:
            for chave, obj in self.ref.items():
                if chave in valor.lower():
                    return obj
            raise Exception('Nome da obrigação não identificado em todos os arquivos, favor selecionar tipo')

    def executar(self):
        # try:   
            if self.matriz.envio_invalido():
                raise Exception ('Insira alguma Matriz')
            elif self.recibos.envio_invalido():
                raise Exception ('Insira algum Recibo')

            declaracao = self.declaracao()
            
            for arquivo in self.recibos.get_caminho():
                declaracao.add_linha(arquivo)

            df = declaracao.gerar_df()

            df_matriz = self.matriz.ler()
            
            if self.valIncrement.get() == True:
                wb_completo = self.matriz.load()
                obj = Incremento(df, df_matriz, wb_completo, declaracao.to_string())
            else:
                obj = Relatorio(df, df_matriz, declaracao.to_string())

            obj.gerar_arquivo()
            
            obj.abrir()
         
        # except (IndexError, TypeError):
        #     messagebox.showerror(title='Aviso', message= 'Erro ao extrair o recibo, confira se a obrigação foi selecionada corretamente. Caso contrário, comunique ao desenvolvedor')
        # except KeyError:
        #     messagebox.showerror(title='Aviso', message= 'Relatório ou Matriz inserido é inválido, certifique-se que inseriu o documento correto')
        # except Exception as error:
        #     messagebox.showerror(title='Aviso', message= error)
       
App()