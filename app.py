import streamlit as st
import pandas as pd

import plotly.express as px

import re
import datetime
import sys

#sap = SapGui()



try:
    @st.cache
    def carregar_data_frames():
        fasor_guarulhos=pd.read_csv(r'C:\Users\b005585\OneDrive - EDP\Documents\db databricks\Fasores\fasor_guarulhos.csv')
        fasor_guarulhos['serial']=fasor_guarulhos['serial'].astype(str)
        fasor_guarulhos.rename(columns={'serial':'medidor'})
        fasor_guarulhos['medidor']= fasor_guarulhos['medidor'].astype(str)
        fasor_guarulhos['numero_instalacao']=fasor_guarulhos['numero'].astype(str)
        fasor_guarulhos.sort_values(by=['date_measur'], inplace=True)
        base_clientes = pd.read_csv(r"C:\Users\b005585\OneDrive - EDP\Documents\db databricks\Base das instalaçoes.csv", dtype=object)
        base_equipamentos = pd.read_csv(r"C:\Users\b005585\OneDrive - EDP\Documents\db databricks\cadastro_equipamentos.csv", dtype=object)
        return fasor_guarulhos, base_clientes, base_equipamentos
        dados = carregar_data_frames()
except:
    st.error("Sem acesso ao sistema da companhia", icon="🚨")

opcao_classe = st.sidebar.selectbox('Selecione uma opção', ['Analise MT', 'Analise BT Ind', 'Analise BT 30/200'])
# fasor_guarulhos=pd.read_csv(r'C:\Users\b005585\OneDrive - EDP\Documents\db databricks\Fasores\fasor_guarulhos.csv')
# fasor_guarulhos['serial']=fasor_guarulhos['serial'].astype(str)
# fasor_guarulhos['numero_instalacao']=fasor_guarulhos['numero_instalacao'].astype(str)
# fasor_guarulhos.sort_values(by=['date_load'], inplace=True)
if opcao_classe == 'Analise MT':
    st.title('ANALISAR BASE DE CONSUMO')
    uploaded_file = st.file_uploader("SELECIONE UM ARQUIVO DE CONSUMO",)

    if uploaded_file is not None:
        bd_consumo= pd.read_excel(uploaded_file,)
        bd_consumo['Instalação']= bd_consumo['Instalação'].astype(str)
        classe = st.selectbox("Selecione uma classe",bd_consumo['Classe'].unique())  
        bd_consumo = bd_consumo[bd_consumo['Classe']==classe]
        colunas1 = bd_consumo.columns
        colunas1= colunas1.drop(['Instalação', 'Classe'])
        opcao = st.selectbox('Selecione um mês',colunas1)
        kwh_min = float(bd_consumo[opcao].min())
        kwh_max = float(bd_consumo[opcao].max())
        def_intervalo = st.slider("Defina um intervalo de consumo kWh/Mês", value= [float(kwh_min), float(kwh_max)],max_value=kwh_max)
        bd_consumo.sort_values(by=[opcao],inplace=True)
        bd_consumo = bd_consumo[ bd_consumo[opcao]>= def_intervalo[0]]
        bd_consumo = bd_consumo[ bd_consumo[opcao]<=def_intervalo[1]]
        texto_expender = "Mostrar Tabela"
        with st.expander(texto_expender,  ):
            st.write('TABELA DE CONSUMO GERAL',bd_consumo[['Instalação',opcao]])
                
        fig = px.bar(bd_consumo.round({'kWh_Mês1':2}), x='Instalação',y=opcao, title='GRÁFICO DE CONSUMO kWh/Mês',color='Instalação')
        st.plotly_chart(fig, use_container_width=True)

        # fig = px.bar(bd_consumo_linha, x='Instalação',y=[['kWh_Mês1','kWh_Mês2',	'kWh_Mês3']],title='GRÁFICO DE CONSUMO kWh/Mês')
        # st.plotly_chart(fig, use_container_width=True)
    
    
    
    st.title('ANALISE DE RELATÓRIO FASORIAL')
    medidor =st.text_input('Digite o numero do medidor')
    analisar = st.button('Analisar')
    if analisar:
        if medidor!="":
            instalação = dados[1][dados[1]['medidor']==medidor]['numero']
            instalação = instalação.tolist()
            try:
                instalação = instalação[0]
                instalação = instalação.lstrip('0')
            except:
                pass
            cliente = dados[1][dados[1]['medidor']==medidor]['nome']
            cliente = cliente.tolist()
            try:
                cliente = cliente[0]
            except:
                pass
            classe = dados[1][dados[1]['medidor']==medidor]['tensao_contratual']
            classe = classe.tolist()
            try:
                classe = classe[0]
            except:
                pass

            cidade =dados[1][dados[1]['medidor']==medidor]['nome_municipio']
            cidade = cidade.tolist()
            try:
                cidade = cidade[0]
            except:
                pass

            try:
                descr_equipamentos=dados[2][dados[2]['ANLAGE']== instalação][['BAUTXT','BAUFORM']]
                for nome_equipamento, info_equipamento in zip(descr_equipamentos['BAUTXT'], descr_equipamentos['BAUFORM']):
                    if (nome_equipamento == "TRANSFORMADOR DE CORRENTE") and ((info_equipamento!= 'TRAFO CORR TRANSIÇÃO (BT)')and (info_equipamento!='TRAFO CORR TRANSIÇÃO (MT)')):
                        equip_tc= info_equipamento
                        equip_tc = equip_tc.split(' ')
                        if len(equip_tc)==4:
                            tensao_tc =equip_tc[2]
                            rtc = equip_tc[3]
                        elif len(equip_tc)==5:
                            tensao_tc =f'{equip_tc[2]}KV'
                            rtc = equip_tc[4]
                        break
                for nome_equipamento, info_equipamento in zip(descr_equipamentos['BAUTXT'], descr_equipamentos['BAUFORM']):
                    if nome_equipamento == "MEDIDOR ELETRÔNICO":
                        equip_med= info_equipamento
                        break
            except:
                pass
            analise_interna = dados[0][dados[0]['serial']==medidor]
                       
            analise_interna_geral = analise_interna[['date_measur','va','vb','vc','ia','ib','ic']]
            analise_interna_i = analise_interna[['date_measur','ia','ib','ic']]
            
            st.write(f'Analise do medidor {medidor} --> {equip_med}')
            st.write(f'Tensão do TC: {tensao_tc}')
            st.write(f'Relação TC: {rtc}')
            st.write(f'Cliente: {cliente}')
            st.write(f'Instalação: {instalação}')
            st.write(f'Classe de tensão: {classe}')
            st.write(f'Cidade: {cidade}')
            with st.expander('Mostrar tabela de fasores'):
                st.write(analise_interna_geral)

            fig_ia = px.line(analise_interna_i,x ='date_measur', y= ['ia','ib','ic'],title="GRÁFICO IA", width=800,height=400)   
            st.plotly_chart(fig_ia,use_container_width=True)
            rtc = "Sem Informação"
            tensao_tc = 'Sem informação'
            # print(f'\n\n\n')
            # print(round(analise_interna.describe(),2))
            # print('-'*50)

