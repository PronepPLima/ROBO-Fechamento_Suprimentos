#10/04/2023
#@PLima

import tkinter as tk
from PIL import Image, ImageTk
from tkinter import filedialog
import pandas as pd
import numpy as py
import warnings
from pandastable import Table
from conect_BD import Conect_bd
import pyautogui
import customtkinter as ctk
import oracledb
import datetime
from cryptography.hazmat.primitives.kdf import pbkdf2

import os


#ignorando alertas exibidos:
warnings.filterwarnings("ignore")

#carregando imagem original:
imagem_original = Image.open("LOGO_PRETA.png")
largura=300
altura=140

#redimencionando a imagem:
imagem_redimencionada = imagem_original.resize((largura,altura))

#Declaracao de variaveis globais:
Unidade = ''
descricoes = []
valores = []

es_df_subtotais =""


# Configurações do banco de dados
server = '10.20.0.129'
database = 'pronep'
username = 'SYSTEM'
password = 'Pronasis1508'

connection = oracledb.connect( user="SYSTEM", password="Pronasis1508", dsn="10.20.0.129/pronep")

cursor = "" 

def criar_pasta_arquivos():
    diretorio = os.path.join(get_exe_directory(), "arquivos")

    if not os.path.exists(diretorio):
        try:
            os.makedirs(diretorio)
            print(f"Pasta '{diretorio}' criada com sucesso!")
            #pyautogui.alert(f"Pasta '{diretorio}' criada com sucesso!") 
        except OSError as e:
            print(f"Erro ao criar pasta '{diretorio}': {e}")
            
            
def get_exe_directory():
    """Retorna o diretório do arquivo executável do programa."""
    return os.path.dirname(os.path.abspath(__file__))



def agora():
    agora = datetime.datetime.now()
    agora = agora.strftime("%d_%m_%Y_%H_%M_%S")
    return str(agora)

def reinicia_label_file_dados():
    label_file_dados.config(text=" ")
    label_file_lista.config(text=" ")


def Seleciona_query_List(opcao):
    print(f'Opcao selecionado: {opcao}')
    
    #TODO: verificar se é RJ e executar a funcao que chama o lista do RJ:
    
    if opcao=="IW_ES":
        print(f'Opcao escolhida: {opcao}')
        fech_esto_V2_lista("IW_PROD_ES")
    elif opcao=="IW_RJ":
        print(f'Opcao escolhida: {opcao}')
        fech_esto_V2_lista("IW_PROD_RJ")
    elif opcao=="IW_SP":
        print(f'Opcao escolhida: {opcao}')
        fech_esto_V2_lista("IW_PROD_SP")
    else:
        print(f'Opcao Invalida!!!')
        
def Seleciona_query_Detalhe(opcao , id_fech):
    print(f'========= Seleciona_query_Detalhe\n')
    
    if opcao=="IW_ES":
        print(f'Opcao escolhida: {opcao}\nId preenchido:{id_fech}')
        fech_esto_V2_detalhado(id_fech , "IW_PROD_ES")
    elif opcao=="IW_RJ":
        print(f'Opcao escolhida: {opcao}\nId preenchido:{id_fech}')
        fech_esto_V2_detalhado(id_fech , "IW_PROD_RJ")
    elif opcao=="IW_SP":
        print(f'Opcao escolhida: {opcao}\nId preenchido:{id_fech}')
        fech_esto_V2_detalhado(id_fech , "IW_PROD_SP")
    else:
        print(f'Opcao Invalida!!!')
    

def fech_esto_V2_lista(esquema):
    global connection
    print(f"\n============================= fech_esto_V2_lista:\n")
    print(f"global connection: {connection}")
    print(f"********** esquema: {esquema}")
    cursor = connection.cursor()
    connection.current_schema = esquema#"IW_PROD_ES"
    print(f"********** connection.current_schema: {connection.current_schema}")
    query = """
                    select 
                        TD_ICW_FECH_EST_C8.ID                                AS ID,                        
                        TO_CHAR(TD_ICW_FECH_EST_C8.DATAT0,'dd/mm/yyyy')      AS DT_INICIO
                    from TD_ICW_FECH_EST_C8 
                    where TD_ICW_FECH_EST_C8.id>0 
                    order by 1 desc
            """
    print(f"Query:\n{query}")
    results = cursor.execute(query)
    print(f"Results:\n{results}")
    data_frame = pd.DataFrame(results, columns=[col[0] for col in results.description])
    #data_frame.to_excel('arquivos\IW_PROD_ES_Lista.xlsx' , index=False)
    cursor.close()
    #editando as datas antes de exibir:
    print(f"data_frame:\n{data_frame}")
    label_file_lista.config(text=data_frame.head(5).to_string(index=False) , justify='center', width=20, height=6)
  
def fech_esto_V2_detalhado(id , esquema):   
    print(f"************* fech_esto_V2_detalhado -> {id}")
    ##retornara um data frame com os dados da query:   
    #
    ##TODO: o problma de nao atualizar parece estar aqui
    #th_es_fech_esto_V2_detalhado = "" 
    #th_es_fech_esto_V2_detalhado = Conect_bd.v2_connection_es_detalhado(id)
    #print(f"\n\n****th_es_fech_esto_V2_detalhado:\n{th_es_fech_esto_V2_detalhado.head(10).to_string(index=False)}")
    #Atualiza_label_file_dados(th_es_fech_esto_V2_detalhado.head(10).to_string(index=False))
    print(f"es_fech_esto_V2_detalhado com o ID: {id}")
    global es_df_subtotais
    
    print(f"global es_df_subtotais:\n{es_df_subtotais}")
    
    global connection
    print(f"global connection: {connection}")
    #pyautogui.alert(f"{connection}")
    print(f"********** esquema: {esquema}")
    #pyautogui.alert(f"{ esquema }")
    cursor = connection.cursor()
    connection.current_schema = esquema#"IW_PROD_ES"
    print(f"**********connection.current_schema: {connection.current_schema}")
    query ="""

                    SELECT
                        --add manualmente a essa query o ID
                        B.ID
                        --, A.CD
                        , F.NAME AS FILIALNAME
                        , TO_CHAR( B.DATAT0, 'DD/MM/YYYY') AS DATAT0
                        , TO_CHAR( (B.DATATF - 1 ) , 'DD/MM/YYYY') AS DATATF
                        , E.Name AS TIPOMATERIAL
                        , A.ScMaterial AS COD_MATERIAL
                        , C.CodeName AS NOME_MATERIAL
                        , D.MU AS UM
                        , (CASE
                            WHEN A.saldo_qt_t0 >= 0
                            THEN A.saldo_qt_t0
                            ELSE 0 END) AS Qtde_Saldo_Inic
                        , (CASE
                            WHEN A.saldo_qt_tf >= 0
                            THEN A.saldo_qt_tf
                            ELSE 0 END) AS Qtde_Saldo_Final
                        , (CASE
                            WHEN (A.ENTR_QT_T0_NF + A.ENTR_QT_T0_INICSD + A.ENTR_QT_T0_DEVOL + A.ENTR_QT_T0_INVENT + A.ENTR_QT_T0_EMP3) >= 0
                            THEN (A.ENTR_QT_T0_NF + A.ENTR_QT_T0_INICSD + A.ENTR_QT_T0_DEVOL + A.ENTR_QT_T0_INVENT + A.ENTR_QT_T0_EMP3)
                            ELSE 0 END ) AS Qtde_E_Total
                        , (CASE
                            WHEN (A.SAIDA_QT_T0_PAC + A.SAIDA_QT_T0_CCNPER + A.SAIDA_QT_T0_CCPERD + A.SAIDA_QT_T0_INVENT + A.SAIDA_QT_T0_DEVOL + A.SAIDA_QT_T0_EMP3) >= 0
                            THEN (A.SAIDA_QT_T0_PAC + A.SAIDA_QT_T0_CCNPER + A.SAIDA_QT_T0_CCPERD + A.SAIDA_QT_T0_INVENT + A.SAIDA_QT_T0_DEVOL + A.SAIDA_QT_T0_EMP3)
                            ELSE 0 END ) AS Qtde_S_TOTAL
                        , (CASE
                            WHEN A.ENTR_QT_T0_NF >= 0
                            THEN A.ENTR_QT_T0_NF
                            ELSE 0 END ) AS Qtde_E_NF
                        , (CASE
                            WHEN A.SAIDA_QT_T0_PAC >= 0
                            THEN A.SAIDA_QT_T0_PAC
                            ELSE 0 END) AS Qtde_S_Pac
                        , (CASE
                            WHEN A.ENTR_QT_T0_INICSD >= 0
                            THEN A.ENTR_QT_T0_INICSD
                            ELSE 0 END) AS Qtde_E_Ini_Sd
                        , (CASE
                            WHEN A.SAIDA_QT_T0_CCNPER >= 0
                            THEN A.SAIDA_QT_T0_CCNPER
                            ELSE 0 END) AS Qtde_S_Outr
                        , (CASE
                            WHEN A.ENTR_QT_T0_DEVOL >= 0
                            THEN A.ENTR_QT_T0_DEVOL
                            ELSE 0 END) AS Qtde_E_Devol
                        , (CASE
                            WHEN A.SAIDA_QT_T0_CCPERD >= 0
                            THEN A.SAIDA_QT_T0_CCPERD
                            ELSE 0 END) AS Qtde_S_Perda
                        , (CASE
                            WHEN A.ENTR_QT_T0_INVENT >= 0
                            THEN A.ENTR_QT_T0_INVENT
                            ELSE 0 END) AS Qtde_E_Invent
                        , (CASE
                            WHEN A.SAIDA_QT_T0_INVENT >= 0
                            THEN A.SAIDA_QT_T0_INVENT
                            ELSE 0 END) AS Qtde_S_Invent
                        , (CASE
                            WHEN A.ENTR_QT_T0_EMP3 >= 0
                            THEN A.ENTR_QT_T0_EMP3
                            ELSE 0 END) AS Qtde_E_3os
                        , (CASE
                            WHEN A.SAIDA_QT_T0_DEVOL >= 0
                            THEN A.SAIDA_QT_T0_DEVOL
                            ELSE 0 END) AS Qtde_S_Devol
                        , (CASE
                            WHEN A.SAIDA_QT_T0_EMP3 >= 0
                            THEN A.SAIDA_QT_T0_EMP3
                            ELSE 0 END) AS Qtde_S_3os
                        , (CASE
                            WHEN A.saldo_vl_t0 >= 0
                            THEN A.saldo_vl_t0
                            ELSE 0 END) AS Valor_Inic
                        , (CASE
                            WHEN A.saldo_vl_tf >= 0
                            THEN A.saldo_vl_tf
                            ELSE 0 END) AS Valor_Final
                        , (CASE
                            WHEN (A.ENTR_VL_INI_SD + A.ENTR_VL_NF_ENTRADA + A.ENTR_VL_INVENT + A.ENTR_VL_DEVOL + A.ENTR_VL_T0_EMP3) >=0
                            THEN (A.ENTR_VL_INI_SD + A.ENTR_VL_NF_ENTRADA + A.ENTR_VL_INVENT + A.ENTR_VL_DEVOL + A.ENTR_VL_T0_EMP3)
                            ELSE 0 END) AS Valor_E_Total
                        , (CASE
                            WHEN (A.VLSAI_QT_T0_PAC + A.VLSAI_QT_T0_CCPERD + A.VLSAI_QT_T0_CCNPER + A.VLSAI_QT_T0_INVENT + A.VLSAI_QT_T0_DEVOL + A.VLSAI_QT_T0_EMP3) >= 0
                            THEN (A.VLSAI_QT_T0_PAC + A.VLSAI_QT_T0_CCPERD + A.VLSAI_QT_T0_CCNPER + A.VLSAI_QT_T0_INVENT + A.VLSAI_QT_T0_DEVOL + A.VLSAI_QT_T0_EMP3)
                            ELSE 0 END) AS Valor_S_Total
                        , (CASE
                            WHEN A.ENTR_VL_NF_ENTRADA >= 0
                            THEN A.ENTR_VL_NF_ENTRADA
                            ELSE 0 END) AS Valor_E_NF
                        , (CASE
                            WHEN A.VLSAI_QT_T0_PAC >= 0
                            THEN A.VLSAI_QT_T0_PAC
                            ELSE 0 END) AS Valor_S_Pac
                        , (CASE
                            WHEN A.ENTR_VL_INI_SD >= 0
                            THEN A.ENTR_VL_INI_SD
                            ELSE 0 END) AS Valor_E_Ini_Sd
                        , (CASE
                            WHEN A.VLSAI_QT_T0_CCNPER >= 0
                            THEN A.VLSAI_QT_T0_CCNPER
                            ELSE 0 END) AS Valor_S_Outr
                        , (CASE
                            WHEN A.ENTR_VL_DEVOL >= 0
                            THEN A.ENTR_VL_DEVOL
                            ELSE 0 END) AS Valor_E_Devol
                        , (CASE
                            WHEN A.VLSAI_QT_T0_CCPERD >= 0
                            THEN A.VLSAI_QT_T0_CCPERD
                            ELSE 0 END) AS Valor_S_Perda
                        , (CASE
                            WHEN A.ENTR_VL_INVENT >= 0
                            THEN A.ENTR_VL_INVENT
                            ELSE 0 END) AS Valor_E_Invent
                        , (CASE
                            WHEN A.VLSAI_QT_T0_INVENT >= 0
                            THEN A.VLSAI_QT_T0_INVENT
                            ELSE 0 END) AS Valor_S_Invent
                        , (CASE
                            WHEN A.ENTR_VL_T0_EMP3 >= 0
                            THEN A.ENTR_VL_T0_EMP3
                            ELSE 0 END) AS Valor_E_3os
                        , (CASE
                            WHEN A.VLSAI_QT_T0_DEVOL >= 0
                            THEN A.VLSAI_QT_T0_DEVOL
                            ELSE 0 END) AS Valor_S_Devol
                        , (CASE
                            WHEN A.VLSAI_QT_T0_EMP3 >= 0
                            THEN A.VLSAI_QT_T0_EMP3
                            ELSE 0 END) AS Valor_S_3os
                    FROM   
                        TD_ICW_FECH_ESTOQ8 A,
                        TD_ICW_FECH_EST_C8 B,
                        SccCode C,
                        MatMaterialType D,
                        SccTable E,
                        MatDispensingArea F
                    WHERE  A.ScMaterial = C.id
                        AND A.ScMaterial = D.ScMaterial
                        AND A.idlot = B.id
                        AND C.idtable = E.ID
                        AND B.ID = :ID
                        AND A.CD = F.ID
                        AND ( A.entr_vl <> 0
                                OR A.saida_vl <> 0
                                OR A.saldo_m <> 0
                                OR A.saldo_qt_tf <> 0
                                OR A.saldo_qt_t0 <> 0
                                OR A.saldo_vl_tf <> 0
                                OR A.saldo_vl_t0 <> 0
                                OR A.entr_qt_t0 <> 0
                                OR A.saida_qt_t0 <> 0
                                OR A.perdas_qt_t0 <> 0
                                OR A.perdas_vl <> 0 )
                        AND ENTR_QT_T0_NF IS NOT NULL
                        AND F.ID = 1 
                    ORDER BY f.name, e.name
        """
    id = id
    print(f'************************** Id: {id}')
    results = cursor.execute(query, [id])
    
    print(f"***** results:\n{results}")
    #pyautogui.alert(f"\n{results}")
    data_frame = pd.DataFrame(results, columns=[col[0] for col in results.description])
    print('\nAqui esta o data frame:\n')
    print(f"{data_frame.info}\n")
    #pyautogui.alert(f"\n{data_frame.info}\n")
    #resultados inseridos num array:
    sub_totais = None
    print(f"***************** Subtotais:\n{sub_totais}")
    print(f"\nAqui esta o SUB_TOTAIS:")
    sub_totais = {
                    (esquema +' - FECHAMENTO'): ('ID: '+ id),
                    'Sub total dietas inicial' : data_frame.query('TIPOMATERIAL == "Dietas"')["VALOR_INIC"].sum(),
                    'Sub total dietas final' : data_frame.query('TIPOMATERIAL == "Dietas"')["VALOR_FINAL"].sum(),
                    'Sub total Mat. Enfermagem inicial' : data_frame.query('TIPOMATERIAL == "Mat. Enfermagem"')["VALOR_INIC"].sum(),
                    'Sub total Mat. Enfermagem final' : data_frame.query('TIPOMATERIAL == "Mat. Enfermagem"')["VALOR_FINAL"].sum(),
                    'Sub total Materiais Diversos inicial' : data_frame.query('TIPOMATERIAL == "Materiais Diversos"')["VALOR_INIC"].sum(),
                    'Sub total Materiais Diversos final' : data_frame.query('TIPOMATERIAL == "Materiais Diversos"')["VALOR_FINAL"].sum(),
                    'Sub total Medicamentos inicial' : data_frame.query('TIPOMATERIAL == "Medicamentos"')["VALOR_INIC"].sum(),
                    'Sub total Medicamentos final' : data_frame.query('TIPOMATERIAL == "Medicamentos"')["VALOR_FINAL"].sum(),
                    'Sub Total estoque inicial' : data_frame["VALOR_INIC"].sum(),
                    'Sub Total estoque final' : data_frame["VALOR_FINAL"].sum(),
                    
                    'Aquisição dietas' : (data_frame.query('TIPOMATERIAL == "Dietas"')["VALOR_E_NF"].sum() + data_frame.query('TIPOMATERIAL == "Dietas"')["VALOR_E_INI_SD"].sum()),
                    'Aquisição Mat. Enfermagem' : (data_frame.query('TIPOMATERIAL == "Mat. Enfermagem"')["VALOR_E_NF"].sum() + data_frame.query('TIPOMATERIAL == "Mat. Enfermagem"')["VALOR_E_INI_SD"].sum()),
                    'Aquisição Materiais Diversos' : (data_frame.query('TIPOMATERIAL == "Materiais Diversos"')["VALOR_E_NF"].sum() + data_frame.query('TIPOMATERIAL == "Materiais Diversos"')["VALOR_E_INI_SD"].sum()),
                    'Aquisição Medicamentos' : (data_frame.query('TIPOMATERIAL == "Medicamentos"')["VALOR_E_NF"].sum() + data_frame.query('TIPOMATERIAL == "Medicamentos"')["VALOR_E_INI_SD"].sum()),
                    'Aquisição estoque' : (data_frame["VALOR_E_NF"].sum() + data_frame["VALOR_E_INI_SD"].sum()),
                    'Devolução dietas ao fornecedor' : data_frame.query('TIPOMATERIAL == "Dietas"')["VALOR_S_DEVOL"].sum(),     
                    'Devolução Mat. Enfermagem ao fornecedor' : data_frame.query('TIPOMATERIAL == "Mat. Enfermagem"')["VALOR_S_DEVOL"].sum(),     
                    'Devolução Materiais Diversos ao fornecedor' : data_frame.query('TIPOMATERIAL == "Materiais Diversos"')["VALOR_S_DEVOL"].sum(),     
                    'Devolução Medicamentos ao fornecedor' : data_frame.query('TIPOMATERIAL == "Medicamentos"')["VALOR_S_DEVOL"].sum(),     
                    'Devolução ao fornecedor' : data_frame["VALOR_S_DEVOL"].sum(),
                    
                    'Envio dietas para paciente ' : (data_frame.query('TIPOMATERIAL == "Dietas"')["VALOR_S_PAC"].sum() - (data_frame.query('TIPOMATERIAL == "Dietas"')["VALOR_S_PAC"].sum()*2)),
                    'Envio Mat. Enfermagem para paciente ' : (data_frame.query('TIPOMATERIAL == "Mat. Enfermagem"')["VALOR_S_PAC"].sum() - (data_frame.query('TIPOMATERIAL == "Mat. Enfermagem"')["VALOR_S_PAC"].sum()*2)),
                    'Envio Materiais Diversos para paciente ' : (data_frame.query('TIPOMATERIAL == "Materiais Diversos"')["VALOR_S_PAC"].sum()-(data_frame.query('TIPOMATERIAL == "Materiais Diversos"')["VALOR_S_PAC"].sum()*2)),
                    'Envio Medicamentos para paciente ' : (data_frame.query('TIPOMATERIAL == "Medicamentos"')["VALOR_S_PAC"].sum()-(data_frame.query('TIPOMATERIAL == "Medicamentos"')["VALOR_S_PAC"].sum()*2)),                    
                    'Envio para paciente ' : (data_frame["VALOR_S_PAC"].sum()-(data_frame["VALOR_S_PAC"].sum()*2)),
                    
                    'Devolução dietas ao paciente' : data_frame.query('TIPOMATERIAL == "Dietas"')["VALOR_E_DEVOL"].sum(),     
                    'Devolução Mat. Enfermagem ao paciente' : data_frame.query('TIPOMATERIAL == "Mat. Enfermagem"')["VALOR_E_DEVOL"].sum(),     
                    'Devolução Materiais Diversos ao paciente' : data_frame.query('TIPOMATERIAL == "Materiais Diversos"')["VALOR_E_DEVOL"].sum(),     
                    'Devolução Medicamentos ao paciente' : data_frame.query('TIPOMATERIAL == "Medicamentos"')["VALOR_E_DEVOL"].sum(),     
                    'Devolução ao paciente' : data_frame["VALOR_E_DEVOL"].sum(),
                    
                    'Ajustes dietas de inventário' : (data_frame.query('TIPOMATERIAL == "Dietas"')["VALOR_E_INVENT"].sum() - data_frame.query('TIPOMATERIAL == "Dietas"')["VALOR_S_INVENT"].sum()),
                    'Ajustes Mat. Enfermagem de inventário' : (data_frame.query('TIPOMATERIAL == "Mat. Enfermagem"')["VALOR_E_INVENT"].sum() - data_frame.query('TIPOMATERIAL == "Mat. Enfermagem"')["VALOR_S_INVENT"].sum()),
                    'Ajustes Materiais Diversos de inventário' : (data_frame.query('TIPOMATERIAL == "Materiais Diversos"')["VALOR_E_INVENT"].sum() - data_frame.query('TIPOMATERIAL == "Materiais Diversos"')["VALOR_S_INVENT"].sum()),
                    'Ajustes Medicamentos de inventário' : (data_frame.query('TIPOMATERIAL == "Medicamentos"')["VALOR_E_INVENT"].sum() - data_frame.query('TIPOMATERIAL == "Medicamentos"')["VALOR_S_INVENT"].sum()),
                    'Ajustes de inventário' : (data_frame["VALOR_E_INVENT"].sum() - data_frame["VALOR_S_INVENT"].sum()),
                                        
                    'Emprestimos dietas' : data_frame.query('TIPOMATERIAL == "Dietas"')["VALOR_E_3OS"].sum(),
                    'Emprestimos Mat. Enfermagem' : data_frame.query('TIPOMATERIAL == "Mat. Enfermagem"')["VALOR_E_3OS"].sum(),
                    'Emprestimos Materiais Diversos' : data_frame.query('TIPOMATERIAL == "Materiais Diversos"')["VALOR_E_3OS"].sum(),
                    'Emprestimos Medicamentos' : data_frame.query('TIPOMATERIAL == "Medicamentos"')["VALOR_E_3OS"].sum(),
                    'Emprestimos ' : data_frame["VALOR_E_3OS"].sum()
                }
    print(f"\nsub_totais:\n{sub_totais}")
    #pyautogui.alert(f"{sub_totais}") 
    #Inicializando o Array para começar do zero:
    descricoes = []
    valores = []
    print(f"\ndescricoes: {descricoes}")
    print(f"valores: {valores}")
    #pyautogui.alert(f"descricoes: {descricoes}\nvalores: {valores}")
    
    #Extraindo Chaves e Valores do Dicionário:
    for chave, valor in sub_totais.items():
        descricoes.append(chave)
        valores.append(valor)
        #pyautogui.alert(f"descricoes: {chave}\nvalores: {valor}")
        
    print(f"\nMARCACAO: \n{data_frame.info}\n")
    #pyautogui.alert(f"\nMARCACAO: \n{data_frame.info}\n")
    label_file_dados.config(text=data_frame.info , justify='right', width=80, height=20)
    
    #Criando o DataFrame:
    df_subtotais = pd.DataFrame({"Descrição": descricoes,"Valor": valores})
    #pyautogui.alert(f"{df_subtotais.info}")
    
    # Nomeando as colunas
    df_subtotais.columns = ["Descrição", "Valor"] 
    print(f"\nes_df_subtotais:\n{df_subtotais.info}")
    #pyautogui.alert(f"{df_subtotais.info}")
    
    #Salvando arquivo xlsx com Esquema:
    local = ('arquivos\\' + esquema + '.xlsx' )
    #local = (esquema + '.xlsx' )
    #pyautogui.alert(f"Local:\n{ local}")
    #Salvando arquivo xlsx com Esquema e ID:
    #local = ('arquivos\\' + esquema + '_' + 'ID_' + id + '.xlsx' )
    
    #Salvando arquivo xlsx com data e hora:
    #local = ('arquivos\\' + esquema + '_' + 'ID_' + id + '_' + agora() + '.xlsx' )
    
    print(f"\nSalvando em:\n{esquema}")
    #pyautogui.alert(f"\nSalvando em:\n{esquema}")
    
    with pd.ExcelWriter(local, engine='openpyxl', mode='w') as writer:
        data_frame.to_excel(writer, sheet_name='Analitico', index=False, header=True)
        df_subtotais.to_excel(writer, sheet_name='Sintetico', index=False, header=False)
        print(f"\nPLANILHA GERADA COM SUCESSO!!!")
        pyautogui.alert(f"          PLANILHA GERADA COM SUCESSO!!!      ")
    
    print(f"\ndata_frame.head(5):\n{data_frame.head(5)}")
    #pyautogui.alert(f"\ndata_frame.head(5):\n{data_frame.head(5)}")
    label_file_dados.config(text=df_subtotais.head(11).to_string(index=False) , justify='right', width=80, height=18)
    cursor.close()
    

        



#============================================================ EXECUCAO ============================================================
if __name__ == "__main__":
    
    try:

        print("\n============================== inicio ========================")
        
        print(f"\n{ get_exe_directory() }")
        
        criar_pasta_arquivos()
        

        #montar interface gráfica:
        root = tk.Tk()
        root.geometry("800x700")
        root.minsize(800,700)
        root.maxsize(800,700)
        #root.maxsize(1000,900)
        root.title("Fechamento Suprimentos 1.0")
        #root.configure(bg="white")
       
        
        
        frame1 = tk.Frame(master=root , width=796, height=180 , border=4 ).place(x=2,y=0)
        
        imagem_tk = ImageTk.PhotoImage(imagem_redimencionada)
        lb_barra_superior = tk.Label(frame1, image=imagem_tk , border =0)
        lb_barra_superior.place(x=250,y=5)
        
        frame2 = tk.Frame(root, width=800 , height=1, background="gray").place(x=2,y=160)
        
        #Zerando opcoes dos radion buttons:
        opcao_var = tk.StringVar()
        opcao_var.set("Nenhuma opcao!")
        
        frame_titulo1 = tk.Frame(root, width=150, height=28, background="gray").place(x=2,y=160)
        label_titulo1 = tk.Label(frame_titulo1, text="Busca Fechamentos:" , background="gray" ,font=("Arial", 10, "bold"),foreground="white").place(x=10 , y=162)
    
        radio_es = tk.Radiobutton(root ,text="IW ES", variable=opcao_var, value='IW_ES' ,command=lambda: [print(f'RadionButton: {opcao_var.get()}') , reinicia_label_file_dados() , Unidade:=opcao_var.get(), print(f'Unidade: {Unidade}')])
        radio_es.place(x=195 , y=190)
        
        radio_rj = tk.Radiobutton(root ,text="IW RJ", variable=opcao_var , value='IW_RJ' ,command=lambda: [print(f'RadionButton: {opcao_var.get()}') , reinicia_label_file_dados() , Unidade:=opcao_var.get(), print(f'Unidade: {Unidade}')])
        radio_rj.place(x=195 , y=210)
    
        radio_sp = tk.Radiobutton(root ,text="IW SP", variable=opcao_var, value='IW_SP' ,command=lambda: [print(f'RadionButton: {opcao_var.get()}') , reinicia_label_file_dados() , Unidade:=opcao_var.get(), print(f'Unidade: {Unidade}')])
        radio_sp.place(x=195 , y=230)
        
        #botao para exibicao de lista de fechamentos:
        bt_list_V2_fech_lista = tk.Button(root, text="Lista Fechamentos V2" , command=lambda: [Unidade:=opcao_var.get() , Seleciona_query_List(Unidade) ])
        bt_list_V2_fech_lista.place(x=310 , y=210)
        
        frame3 = tk.Frame(root , width=146,height=96, background="light gray").place(x=520 , y=170)
        label_file_lista = tk.Label(frame3,text='',border =0 , width=20 , height=6)
        label_file_lista.place(x=522 , y=172)
        
        frame4 = tk.Frame(root, width=800 , height=1, background="gray").place(x=2,y=272)
        
        #frame_titulo2 = tk.Frame(root, width=160 , height=2, background="gray" ).place(x=12,y=300)
        #label_titulo2 = tk.Label(frame_titulo2, text="Busca materiais/medicamentos:")
        #label_titulo2.place(x=10 , y=274)
        
        frame_titulo2 = tk.Frame(root, width=220, height=28, background="gray").place(x=2,y=274)
        label_titulo2 = tk.Label(frame_titulo2, text="Busca materiais/medicamentos:" , background="gray" ,font=("Arial", 10, "bold"),foreground="white").place(x=10 , y=274)

        #botao para V2_fech_esto_RJ
        label_campo_entrada = tk.Label(root, text="Digite o ID:").place(x=195 , y=310)
        campo_entrada = tk.Entry(root, width=3)
        campo_entrada.place(x=270 , y=310)
                
        bt_V2_fech_esto_detalh = tk.Button(root, text="Buscar Fechamento V2" , command=lambda: [ Unidade:=opcao_var.get() , Seleciona_query_Detalhe(Unidade , campo_entrada.get())])
        bt_V2_fech_esto_detalh.place(x=310,y=310)
        
        #label_file_dados = tk.Label(root,text='',border =0 , width=80 , height=12)
        #label_file_dados.place(x=290 , y=340)

        
        # Criar Label para exibir a tabela
        label_file_dados = tk.Label(root, text="")
        label_file_dados.place(x=290 , y=340)
               
        label_rodape = tk.Label(root,text='Uso exclusivo da coordenação de suprimentos/farmácia Pronep.',border =0)
        label_rodape.place(x=453 , y=678)
       
        root.mainloop()
        print("\n============================== fim ========================")

    
    except Exception as err:
        print(f"Erro Inexperado: {err=}, \n{type(err)=}")
        pyautogui.alert(f"Erro Inexperado: {err=}, \n{type(err)=}")