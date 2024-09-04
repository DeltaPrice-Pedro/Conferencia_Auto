from tkinter import *
from tkinter import messagebox
from tkinter.filedialog import askopenfilenames, askopenfilename, asksaveasfilename
import tabula as tb
import pandas as pd
from unidecode import unidecode
import string
import os

from datetime import *
from json import loads
from xlsxwriter import *
import locale

locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')

window = Tk()

class Arquivo:
    def __init__(self):
        self.tipos_validos = ''

    def validar_entrada(self, caminhos):
        text_caminhos = ''
        for caminho in caminhos:
            if any(c not in string.ascii_letters for c in caminho):
                caminho = self.formato_ascii(caminho)

            ultima_barra = caminho.rfind('/')
            self.__tipo(caminho)
            text_caminhos = text_caminhos + '\n' + caminho[ultima_barra+1:]

        return text_caminhos

    def __tipo(self, caminho):
        if caminho[len(caminho) -3 :] != self.tipos_validos:
            ultima_barra = caminho.rfind('/')
            raise Exception(
                f'Formato inválido do arquivo: {caminho[ultima_barra+1:]}')

    def formato_ascii(self, caminho):
        caminho_uni = unidecode(caminho)
        os.rename(caminho, caminho_uni)
        return caminho_uni
    
    def get_caminho(self):
        return self.caminho

class Matriz(Arquivo):
    def __init__(self):
        super().__init__()
        self.tipos_validos = 'lsx'
        self.caminho = ''

    def inserir(self, label):
        try:
            caminhos = [askopenfilename()]

            if caminhos == '':
                raise ValueError('Operação cancelada')

            label['text'] = self.validar_entrada(caminhos)

            self.caminho = caminhos

        except ValueError:
            messagebox.showerror(title='Aviso', message= 'Operação cancelada')
        except Exception as error:
            messagebox.showerror(title='Aviso', message= error)

    # def validar(self, df):
    #     cnpjs_matriz = self.ler()
    #     qnt_excluidos = 0

    #     for index, row in df.iterrows():
    #         if row['CNPJ'] not in cnpjs_matriz:
    #             qnt_excluidos = qnt_excluidos + 1

    #     return qnt_excluidos

    def ler(self):
        return pd.read_excel(self.caminho[0], na_filter=False, usecols='A:B')
    
    def cnpjs(self, arquivo):
        return arquivo.loc[arquivo['CNPJ'] != '']

class Recibo(Arquivo):
    def __init__(self):
        super().__init__()
        self.tipos_validos = 'pdf'
        self.caminho = []

    def inserir(self, label):
        try:
            caminhos = askopenfilenames()

            if caminhos == []:
                raise ValueError('Operação cancelada')

            label['text'] = self.validar_entrada(caminhos)

            self.caminho = caminhos

        except ValueError:
            messagebox.showerror(title='Aviso', message= 'Operação cancelada')
        except Exception as error:
            messagebox.showerror(title='Aviso', message= error)

class Writer:
    def __init__(self, df, df_matriz, titulo):
        self.df = df
        self.df_matriz = df_matriz
        self.LIN_INDEX = 6
        self.lin_data = self.LIN_INDEX + 1
        self.titulo = titulo

    def abrir(self):
        file = asksaveasfilename(title='Favor selecionar a pasta onde será salvo', filetypes=((".xlsx","*.xlsx"),))

        self.gerar_arquivo(file)

        messagebox.showinfo(title='Aviso', message='Abrindo o arquivo gerado!')

        os.startfile(file+'.xlsx')

    def gerar_arquivo(self, nome_arq):
        self.wb = Workbook(nome_arq + '.xlsx')
        self.ws = self.wb.add_worksheet('Sheet1')
        self.ws.set_column('A:A', 40)
        self.ws.set_column('B:B', 20)
        self.ws.set_column('C:C', 10)
        self.ws.set_column('D:D', 20)
        self.ws.set_column('E:E', 15)
        self.ws.set_column('F:F', 20)

        self.cabecalho()
        self.table_ref()
        self.preencher_fields()
        self.preencher_matriz()
        
        excluidos = self.preencher_data()

        if excluidos != []:
            messagebox.showinfo(title='Aviso', message= f"Os recibos das empresas: - {excluidos} - foram marcados por não constarem na matriz")

        self.wb.close()

    def cabecalho(self):
        self.ws.write(0,0,f'RELATÓRIO DE CONFERÊNCIA {self.titulo}',\
            self.wb.add_format({'bold': True, 'font_size': 26}))

        self.ws.write(2,0,'Competência',\
            self.wb.add_format({'bold':True,'align':'right','font_size': 16}))
        self.ws.write(2,1, self.data_confe())

        self.ws.write(3,0,'Data Entrega',\
            self.wb.add_format({'bold':True,'align':'right','font_size': 16}))
        self.ws.write(3,1, datetime.now().strftime("%d/%m/%Y"))

    def table_ref(self):
        tam_df = len(self.df_matriz)+7
        ref = {
            'Obrigadas:': f'=COUNTA($B8:$B{tam_df})',
            'Entregues:': f'=COUNTIF($D8:$D{tam_df}, "ENVIADO")',
            'Não Entregues:': '=$E3 - $E4'
        }
        for index, text in enumerate(ref.items()):
            self.ws.write(index+2,3, text[0],\
                self.wb.add_format({'bold':True,'border':1,'align':'right'}))
            
            self.ws.write(index+2,4, text[1],\
                self.wb.add_format({'border':1,'align':'center'}))

    def preencher_fields(self):
        for index, columns in enumerate(list(self.df.columns)):
            self.ws.write(self.LIN_INDEX, index, columns,\
                self.wb.add_format({'bold':True,'top':2, 'bg_color':'#a7b8ab','underline':True, 'align':'center'}))

    def preencher_matriz(self):
        for index, row in self.df_matriz.iterrows():
            #Adciona Nome e CNPJ apenas
            for col_index, valor in enumerate(row):
                self.ws.write(index + self.lin_data, col_index, valor,\
                    self.wb.add_format({'border':3, 'align':'center'}))
            
            #Termina o df com espaços vazios
            self.espacos_vazios(index)

    def espacos_vazios(self, index):
        diferenca = 1
        col_index = 2
        for diferenca in range(len(self.df)+1):
            self.ws.write(index + self.lin_data, col_index + diferenca, '',\
                self.wb.add_format({'border':3}))

    def preencher_data(self):
        lista_excluidos = []
        cnpj_matriz = Matriz().cnpjs(self.df_matriz)
        for index, row in self.df.iterrows():
            if row['CNPJ'] not in cnpj_matriz.values:
                formato = self.wb.add_format({'border':3, 'align':'center', 'bg_color':'yellow'})
                lista_excluidos.append(row['Nome'])
            else:
                formato = self.wb.add_format({'border':3, 'align':'center'})
                
            for col_index, valor in enumerate(row):
                self.ws.write(index + self.lin_data, col_index, valor, formato)

        return lista_excluidos

    def data_confe(self):
        data = f'{datetime.now().month - 1}/{datetime.now().year}'
        data_format = datetime.strptime(data, '%m/%Y')
        return data_format.strftime("%B/%Y".capitalize())

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
        self.tabela = tb.read_pdf(arquivo, pages= 1, stream= True,\
                        relative_area=True, area= [10,0,59,68])[0]
        
        ##CNPJ
        self.cnpj.append(self.tabela.iloc[0,1])

        ##Nome Emp
        self.nome_emp.append(self.tabela.iloc[1,0]\
            .replace('Nome/Razão Social: ',''))

        ##Ref.
        self.referencia.append(self.tabela.iloc[3,0]\
            .replace('Referência: ','')\
                    .replace(' No Protocolo:',''))

        ##Data e Hora
        col_dthr = self.tabela.iloc[4,0].replace('Data/Hora de Entrega: ','')\
                        .replace(' Regime de Tributação:','')

        self.data.append(col_dthr[:10])

        self.hora.append(col_dthr[10:])

        ##Serviços declarados.
        self.servicos.append(self.tabela.iloc[15,0]\
            .replace('Total de Serviços Declarados: ','')\
                .replace('Base de Cálculo S/ Ret',''))

    def gerar_df(self):
        return pd.DataFrame({
            'Nome': self.nome_emp, 
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
            #'Num. Domínio': self.num_dom,
            'Nome': self.nome_emp, 
            'CNPJ': self.cnpj, 
            'Referência': self.referencia,
            'Situação': self.situacao, 
            'Data Entrega': self.data,
            'Hora Entrega': self.hora,
            })

class App:
    def __init__(self):
        self.window = window
        self.recibos = Recibo()
        self.matriz = Matriz()
        self.tela()
        self.index()
        window.mainloop()

    def tela(self):
        self.window.configure(background='darkblue')
        self.window.resizable(False,False)
        self.window.geometry('860x500')
        #self.window.iconbitmap('Z:\\18 - PROGRAMAS DELTA\\code\\imgs\\delta-icon.ico')
        self.window.title('Conversor de Extrato')

    def index(self):
        self.index = Frame(self.window, bd=4, bg='lightblue')
        self.index.place(relx=0.05,rely=0.05,relwidth=0.9,relheight=0.9)

        #Titulo
        Label(self.index, text='Conferência Automática', background='lightblue', font=('arial',30,'bold')).place(relx=0.23,rely=0.1,relheight=0.15)

        #Logo
        # self.logo = PhotoImage(file='Z:\\18 - PROGRAMAS DELTA\\code\\imgs\\deltaprice-hori.png')
        
        # self.logo = self.logo.subsample(4,4)
        
        # Label(self.window, image=self.logo, background='lightblue')\
        #     .place(relx=0.175,rely=0.1,relwidth=0.7,relheight=0.2)

        #Labels e Entrys
        ###########Matriz
        Label(self.index, text='Insira aqui a Matriz/Referência:',\
            background='lightblue', font=(10))\
                .place(relx=0.15,rely=0.25)

        self.nome_Mat = ''
        self.matLabel = Label(self.index)
        self.matLabel.config(font=("Arial", 8, 'bold italic'))
        self.matLabel.place(relx=0.21,rely=0.32,relwidth=0.7, relheight=0.055)
        
        Button(self.index, text='Enviar',\
            command = lambda: self.matriz.inserir(self.matLabel))\
                .place(relx=0.15,rely=0.32,relwidth=0.06,relheight=0.055)

        ###########Arquivo
        Label(self.index, text='Insira aqui os Recibos:',\
            background='lightblue', font=(10))\
                .place(relx=0.15,rely=0.4)

        self.nome_arq = ''
        self.arqLabel = Label(self.index)
        self.arqLabel.config(font=("Arial", 8, 'bold italic'))
        self.arqLabel.place(relx=0.21,rely=0.47,relwidth=0.7, relheight=0.2)
        
        Button(self.index, text='Enviar',\
            command = lambda: self.recibos.inserir(self.arqLabel))\
                .place(relx=0.15,rely=0.47,relwidth=0.06,relheight=0.055)

        ###########EFD
        Label(self.index, text='Caso o nome da obrigação assesória não constar no nome do arquivo',\
            background='lightblue', font=("Arial", 12, 'bold italic'))\
                .place(relx=0.15,rely=0.7)

        Label(self.index, text='Escolha a obrigação:',\
            background='lightblue', font=(10))\
                .place(relx=0.15,rely=0.75)
        
        self.declaracaoEntry = StringVar()

        self.declaracaoEntryOpt = ["DES", "REINF"]

        self.declaracaoEntry.set('Escolha aqui')

        self.popup = OptionMenu(self.index, self.declaracaoEntry, *self.declaracaoEntryOpt)\
            .place(relx=0.375,rely=0.75,relwidth=0.2,relheight=0.06)
        
        #Botão enviar
        Button(self.index, text='Gerar Conferencia',\
            command= lambda: self.executar())\
                .place(relx=0.65,rely=0.8,relwidth=0.25,relheight=0.12)

    def definir_declaracao(self):
        declaracao_selecionado = self.declaracaoEntry.get()
        nome_arq = self.arqLabel['text']

        if declaracao_selecionado == 'DES' or\
            'des' in nome_arq.lower():
            return Des()
        elif declaracao_selecionado == 'REINF' or\
            'reinf' in nome_arq.lower():
            return Reinf()
        raise Exception('Declaração inválida, favor selecioná-lo')

    def executar(self):
        #try:       
            cam_matriz = self.matriz.get_caminho()
            cam_recibo = self.recibos.get_caminho()

            if cam_recibo == (''):
                raise Exception ('Insira algum Recibo')
            elif cam_matriz == '':
                raise Exception ('Insira alguma Matriz')

            declaracao = self.definir_declaracao()

            for arquivo in cam_recibo:
                declaracao.add_linha(arquivo)

            df = declaracao.gerar_df()

            df_matriz = self.matriz.ler()
            
            Writer(df, df_matriz, declaracao.to_string()).abrir()
         
        # except PermissionError:
        #     messagebox.showerror(title='Aviso', message= 'Feche o arquivo gerado antes de criar outro')
        # except UnboundLocalError:
        #     messagebox.showerror(title='Aviso', message= 'Arquivo não compativel a esse banco')
        # except subprocess.CalledProcessError:
        #     messagebox.showerror(title='Aviso', message= "Erro ao extrair a tabela, confira se o banco foi selecionado corretamente. Caso contrário, comunique o desenvolvedor")
        # except Exception as error:
        #     messagebox.showerror(title='Aviso', message= error)
       
App()