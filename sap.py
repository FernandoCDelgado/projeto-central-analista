import subprocess
import sys
import time
import pandas as pd
import win32com.client
import os
import pythoncom
pythoncom.CoInitialize()
from win32com.client import GetObject


#Criando a classe sap
class SapGui():
    #Função para conectar sap
    def connection_sap(self):
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        application = SapGuiAuto.GetScriptingEngine
        self.connection = application.Children(0)
        self.session = self.connection.Children(0)

        
    #Função para abrir e fazer login no sap
    def sap_login(self):
        try:
            path = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
            subprocess.Popen(path)
            time.sleep(3)
            SapGuiAuto = win32com.client.GetObject('SAPGUI')
            if not type(SapGuiAuto) == win32com.client.CDispatch:
                return

            application = SapGuiAuto.GetScriptingEngine
            if not type(application) == win32com.client.CDispatch:
                SapGuiAuto = None
                return
            connection = application.OpenConnection("PRODUÇÃO CCS ( EP1 ) - EDP SP", True)

            if not type(connection) == win32com.client.CDispatch:
                application = None
                SapGuiAuto = None
                return

            self.session = connection.Children(0)
            if not type(self.session) == win32com.client.CDispatch:
                    connection = None
                    application = None
                    SapGuiAuto = None
                    return

            self.session.findById("wnd[0]/usr/txtRSYST-BNAME").text = "B005585"
            self.session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = "Nov@2022"
            self.session.findById("wnd[0]").sendVKey(0)
            self.session.findById("wnd[0]").minimize()
        except:
            print(sys.exc_info())

        
        # finally:
        #     session = None
        #     connection = None
        #     application = None
        #     SapGuiAuto = None
    def fechar_sap(self):
        self.connection_sap()
        self.connection.CloseSession('ses[0]') 

    def consultar_consumo(self, instalacao):
        time.sleep(3)
        self.session.findById("wnd[0]/tbar[0]/okcd").text = "es32"
        self.session.findById("wnd[0]").sendVKey (0)
        self.session.findById("wnd[0]/usr/ctxtEANLD-ANLAGE").text = instalacao
        self.session.findById("wnd[0]").sendVKey (0)
        status_instalação = self.session.findById("wnd[0]/usr/txtEANLD-DISCSTAT").text
        cliente = self.session.findById("wnd[0]/usr/txtEANLD-PARTTEXT").text
        self.tensao_contratada = self.session.findById("wnd[0]/usr/txtEANLD-SPEBENETXT").text
        text = self.session.findById("wnd[0]/usr/txtEANLD-ALARTSTTXT").text
        contrato = self.session.findById("wnd[0]/usr/txtEANLD-VERTRAG").text
        self.session.findById("wnd[0]/tbar[0]/btn[3]").press()
        self.session.findById("wnd[0]/tbar[0]/btn[3]").press()
        self.session.findById("wnd[0]/tbar[0]/okcd").text = "/nzccs_consumo"
        self.session.findById("wnd[0]").sendVKey (0)
        self.session.findById("wnd[0]/usr/ctxtPC_VERTR-LOW").text = contrato
        self.session.findById("wnd[0]/usr/ctxtPC_VERTR-LOW").setFocus
        self.session.findById("wnd[0]/tbar[1]/btn[8]").press()
        self.session.findById("wnd[0]/tbar[1]/btn[45]").press()
        self.session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").select()
        self.session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").setFocus
        self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
        self.session.findById("wnd[0]/tbar[0]/btn[3]").press()
        self.session.findById("wnd[0]/tbar[0]/btn[3]").press()
        #Gera um dataframe com os dados da clipboard
        self.dados = pd.read_clipboard(sep="|", skiprows=range(0,6), on_bad_lines= 'skip', dtype=object, encoding= "UTF-8" )
        #Tira os espaços vazios das colunas
        #Cria um nova coluna com o texto em formato de moeda com o valor pago em reais 
        self.dados.columns = self.dados.columns.str.strip()
        self.dados.columns
        #Retira as colunas desnecessarias do dataframe
        try:
            self.dados.drop(['Unnamed: 0', '', 'Compl', 'E', 'Extrato','Média de C','Tipo de Fa',
            'Unnamed: 17', '','.1', '', '', '.1', '.2', '.3', 'Unnamed: 25'], axis=1, inplace=True)
        except:
            self.dados.drop(['Unnamed: 0', '', 'Compl', 'E', 'Extrato',
                'Média de C',
            'Tipo de Fa', '', '', '.1', '',
            '', '.1', '.2', '.3', 'Unnamed: 25'], axis=1, inplace=True)
        #Exclui as linhas com self.dados nulos de acordo com o critério adotado
        self.dados.dropna(subset=['Mês/Ano', 'Leitura'], inplace = True)
        #Retira os espaços vazios de todas as linha da coluna Função
        self.dados['Função'] = self.dados['Função'].str.strip()
        #Retira todos os espaços vazios e transforma os self.dados em float na coluna Cons/Dem
        self.dados['Cons/Dem'] = self.dados['Cons/Dem'].str.strip()
        self.dados['Cons/Dem']= self.dados['Cons/Dem'].apply(lambda x: float(x.replace('.','').replace(',','.')))
        
        #Cria um nova coluna com o texto em formato de moeda com o valor pago em reais 
        lista_coluna = self.dados.columns.tolist()
        if 'Total da F' in lista_coluna:
            list=[]
            for valor in self.dados['Total da F']:
                valor = f'R${valor}'
                list.append(valor)
            self.dados['Fatura']= list
        elif "Montante" in lista_coluna:
            list=[]
            for valor in self.dados['Montante']:
                valor = f'R${valor}'
                list.append(valor)
            self.dados['Fatura']= list
        #Trata a coluna "Função" para interarmos sobre a mesma e grava no dataframe "dados_canais"
        list_função = self.dados['Função'].tolist()
        
        if "08" in list_função:
            self.dados_canais = self.dados.loc[self.dados['Função']=="08"]
        elif "03" in list_função:
            self.dados_canais = self.dados.loc[self.dados['Função']=="03"]

        #Trata a coluna "Mês/Ano" para pegar somente o valor referente ao ano
        self.dados_canais["Ano"]= self.dados_canais['Mês/Ano'].str[3:]

        #Classifica os dados em ordem decrescente aplicando na coluna "Ano" e "Mês/Ano"
        self.dados_canais = self.dados_canais.sort_values(by=(["Ano", 'Mês/Ano'])) 
        lista_coluna = self.dados_canais.columns.tolist()
        if 'Total da F' in lista_coluna:
            self.dados_canais['Total da F']= self.dados_canais['Total da F'].apply(lambda x: float(x.replace('.','').replace(',','.').replace('-','')))
            estatistica = self.dados_canais.agg({ 'Cons/Dem':['max','mean', 'std', 'min'],
                                'Total da F':['max','mean', 'std', 'min'],
        })
        elif "Montante" in lista_coluna:
            self.dados_canais["Montante"]= self.dados_canais["Montante"].apply(lambda x: float(x.replace('.','').replace(',','.').replace('-','')))
            estatistica =self.dados_canais.agg({ 'Cons/Dem':['max','mean', 'std', 'min'],
                                'Montante':['max','mean', 'std', 'min'],
                            })
       
        return [self.dados, self.dados_canais, status_instalação, cliente,]
        



    def consultar_massa(self, instalação):    
        self.session.findById("wnd[0]/tbar[0]/okcd").text = "es32"
        self.session.findById("wnd[0]").sendVKey (0)
        self.session.findById("wnd[0]/usr/ctxtEANLD-ANLAGE").text = instalação
        self.session.findById("wnd[0]").sendVKey (0)
        status_instalação = self.session.findById("wnd[0]/usr/txtEANLD-DISCSTAT").text
        cliente = self.session.findById("wnd[0]/usr/txtEANLD-PARTTEXT").text
        self.tensao_contratada = self.session.findById("wnd[0]/usr/txtEANLD-SPEBENETXT").text
        text = self.session.findById("wnd[0]/usr/txtEANLD-ALARTSTTXT").text
        contrato = self.session.findById("wnd[0]/usr/txtEANLD-VERTRAG").text
        if contrato != "":
            self.session.findById("wnd[0]/tbar[0]/btn[3]").press()
            self.session.findById("wnd[0]/tbar[0]/btn[3]").press()
            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/nzccs_consumo"
            self.session.findById("wnd[0]").sendVKey (0)
            self.session.findById("wnd[0]/usr/ctxtPC_VERTR-LOW").text = contrato
            self.session.findById("wnd[0]/usr/ctxtPC_VERTR-LOW").setFocus
            try:
                self.session.findById("wnd[0]/tbar[1]/btn[8]").press()
                self.session.findById("wnd[0]/tbar[1]/btn[45]").press()
                self.session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").select()
                self.session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").setFocus
                self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
                self.session.findById("wnd[0]/tbar[0]/btn[3]").press()
                self.session.findById("wnd[0]/tbar[0]/btn[3]").press()
            except:
                self.session.findById("wnd[0]/tbar[0]/btn[3]").press()
                contrato= ""
            #Gera um dataframe com os dados da clipboard
            self.dados = pd.read_clipboard(sep="|", skiprows=range(0,6), on_bad_lines= 'skip', dtype=object, encoding= "UTF-8" )
            #Tira os espaços vazios das colunas
            self.dados.columns = self.dados.columns.str.strip()
            #Exclui as linhas com self.dados nulos de acordo com o critério adotado
            self.dados.dropna(subset=['Mês/Ano', 'Leitura'], inplace = True)
            #Retira os espaços vazios de todas as linha da coluna Função
            self.dados['Função'] = self.dados['Função'].str.strip()
            #Retira todos os espaços vazios e transforma os self.dados em float na coluna Cons/Dem
            self.dados['Cons/Dem'] = self.dados['Cons/Dem'].str.strip()
            self.dados['Cons/Dem']= self.dados['Cons/Dem'].apply(lambda x: float(x.replace('.','').replace(',','.')))
            #Laço para verificar o valor do canal e definir se é MT ou BT
            verificador = self.dados["Função"]
            verificador = verificador.tolist()
            if "I3" in verificador:
                self.classe_tensao= "BT MMGD"
                self.dados_canais = self.dados.loc[self.dados['Função']=="03"]
            elif "03" in verificador:
                self.classe_tensao= "BT"
                self.dados_canais = self.dados.loc[self.dados['Função']=="03"]
            elif "04" in verificador:
                self.classe_tensao= "MT"
                self.dados_canais = self.dados.loc[self.dados['Função']=="08"]
            
            # for canal in self.dados["Função"]:
                
                
            #     if canal == "03":
            #         self.dados_canais = self.dados.loc[self.dados['Função']=="03"]
            #         self.classe_tensao = "BT"
            #         break
            #     elif canal == "04":
            #         self.dados_canais = self.dados.loc[self.dados['Função']=="08"]
            #         self.classe_tensao="MT"
            #         break
            #Carrega a variavel com o valor maximo de consumo do dataframe
            valor_max = self.dados_canais.loc[self.dados_canais['Cons/Dem']==self.dados_canais['Cons/Dem'].max()]
            #Carrega a variável com o valor médio de consumo do dataframe
            valor_medio = round(self.dados_canais['Cons/Dem'].mean(),2)
            #Carrega o valor da variavel com o valor minimo
            valor_minimo = self.dados.loc[self.dados['Cons/Dem']==self.dados_canais['Cons/Dem'].min()]
            #Carrega a variável com o valor do desvio padrão
            valor_std = round(self.dados_canais['Cons/Dem'].std())
        
            return [self.dados, contrato, self.classe_tensao] 
        else: 
            self.session.findById("wnd[0]/tbar[0]/btn[3]").press()
            return self.dados, contrato

    def es32(self):    
        time.sleep(3)
        self.session.findById("wnd[0]/tbar[0]/okcd").text = "es32"

    def analise_massa(self, instalacao):
        # time.sleep(3)
        # self.session.findById("wnd[0]/tbar[0]/okcd").text = "es32"
        self.session.findById("wnd[0]").sendVKey (0)
        self.session.findById("wnd[0]/usr/ctxtEANLD-ANLAGE").text = instalacao
        self.session.findById("wnd[0]").sendVKey (0)
        status_instalação = self.session.findById("wnd[0]/usr/txtEANLD-DISCSTAT").text
        cliente = self.session.findById("wnd[0]/usr/txtEANLD-PARTTEXT").text
        self.tensao_contratada = self.session.findById("wnd[0]/usr/txtEANLD-SPEBENETXT").text
        text = self.session.findById("wnd[0]/usr/txtEANLD-ALARTSTTXT").text
        contrato = self.session.findById("wnd[0]/usr/txtEANLD-VERTRAG").text
        if contrato!= '':
            self.session.findById("wnd[0]/tbar[0]/btn[3]").press()
            self.session.findById("wnd[0]/tbar[0]/btn[3]").press()
            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/nzccs_consumo"
            self.session.findById("wnd[0]").sendVKey (0)
            self.session.findById("wnd[0]/usr/ctxtPC_VERTR-LOW").text = contrato
            self.session.findById("wnd[0]/usr/ctxtPC_VERTR-LOW").setFocus
            self.session.findById("wnd[0]/tbar[1]/btn[8]").press()
            self.session.findById("wnd[0]/tbar[1]/btn[45]").press()
            self.session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").select()
            self.session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").setFocus
            self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
            self.session.findById("wnd[0]/tbar[0]/btn[3]").press()
            #self.session.findById("wnd[0]/tbar[0]/btn[3]").press()
            #Gera um dataframe com os dados da clipboard
            self.dados = pd.read_clipboard(sep="|", skiprows=range(0,6), on_bad_lines= 'skip', dtype=object, encoding= "UTF-8" )
            #Tira os espaços vazios das colunas
            #Cria um nova coluna com o texto em formato de moeda com o valor pago em reais 
            self.dados.columns = self.dados.columns.str.strip()
            self.dados.columns
            #Retira as colunas desnecessarias do dataframe
            try:
                self.dados.drop(['Unnamed: 0', '', 'Compl', 'E', 'Extrato','Média de C','Tipo de Fa',
                'Unnamed: 17', '','.1', '', '', '.1', '.2', '.3', 'Unnamed: 25'], axis=1, inplace=True)
            except:
                self.dados.drop(['Unnamed: 0', '', 'Compl', 'E', 'Extrato',
                    'Média de C',
                'Tipo de Fa', '', '', '.1', '',
                '', '.1', '.2', '.3', 'Unnamed: 25'], axis=1, inplace=True)
            #Exclui as linhas com self.dados nulos de acordo com o critério adotado
            self.dados.dropna(subset=['Mês/Ano', 'Leitura'], inplace = True)
            #Retira os espaços vazios de todas as linha da coluna Função
            self.dados['Função'] = self.dados['Função'].str.strip()
            #Retira todos os espaços vazios e transforma os self.dados em float na coluna Cons/Dem
            self.dados['Cons/Dem'] = self.dados['Cons/Dem'].str.strip()
            self.dados['Cons/Dem']= self.dados['Cons/Dem'].apply(lambda x: float(x.replace('.','').replace(',','.')))
            
            #Cria um nova coluna com o texto em formato de moeda com o valor pago em reais 
            lista_coluna = self.dados.columns.tolist()
            if 'Total da F' in lista_coluna:
                list=[]
                for valor in self.dados['Total da F']:
                    valor = f'R${valor}'
                    list.append(valor)
                self.dados['Fatura']= list
            elif "Montante" in lista_coluna:
                list=[]
                for valor in self.dados['Montante']:
                    valor = f'R${valor}'
                    list.append(valor)
                self.dados['Fatura']= list
            #Trata a coluna "Função" para interarmos sobre a mesma e grava no dataframe "dados_canais"
            list_função = self.dados['Função'].tolist()
            
            if "08" in list_função:
                self.dados_canais = self.dados.loc[self.dados['Função']=="08"]
            elif "03" in list_função:
                self.dados_canais = self.dados.loc[self.dados['Função']=="03"]

            #Trata a coluna "Mês/Ano" para pegar somente o valor referente ao ano
            self.dados_canais["Ano"]= self.dados_canais['Mês/Ano'].str[3:]

            #Classifica os dados em ordem decrescente aplicando na coluna "Ano" e "Mês/Ano"
            self.dados_canais = self.dados_canais.sort_values(by=(["Ano", 'Mês/Ano'])) 
            lista_coluna = self.dados_canais.columns.tolist()
            if 'Total da F' in lista_coluna:
                self.dados_canais['Total da F']= self.dados_canais['Total da F'].apply(lambda x: float(x.replace('.','').replace(',','.').replace('-','')))
                estatistica = self.dados_canais.agg({ 'Cons/Dem':['max','mean', 'std', 'min'],
                                    'Total da F':['max','mean', 'std', 'min'],
            })
            elif "Montante" in lista_coluna:
                self.dados_canais["Montante"]= self.dados_canais["Montante"].apply(lambda x: float(x.replace('.','').replace(',','.').replace('-','')))
                estatistica =self.dados_canais.agg({ 'Cons/Dem':['max','mean', 'std', 'min'],
                                    'Montante':['max','mean', 'std', 'min'],
                                })
        
            return [self.dados, self.dados_canais, status_instalação, cliente,]
        else:
            self.session.findById("wnd[0]/tbar[0]/btn[3]").press()