#10/04/2023
#@PLima


import tkinter as tk
from PIL import Image, ImageTk
from tkinter import filedialog
import pandas as pd
import numpy as py
import warnings
from pandastable import Table
#from conect_BD import v2_connection_rj_detalhado
from conect_BD import Conect_bd
import threading
import pyautogui

#ignorando alertas exibidos:
warnings.filterwarnings("ignore")

#carregando imagem original:
imagem_original = Image.open("LOGO_PRETA.png")
largura=300
altura=140

#redimencionando a imagem:
imagem_redimencionada = imagem_original.resize((largura,altura))


Unidade = ''

def Seleciona_query_List(opcao):
    print(f'Opcao selecionado: {opcao}')
    
    #TODO: verificar se é RJ e executar a funcao que chama o lista do RJ:
    
    if opcao=="IW_ES":
        print(f'Opcao escolhida: {opcao}')
        es_fech_esto_V2_lista()
    elif opcao=="IW_RJ":
        print(f'Opcao escolhida: {opcao}')
        rj_fech_esto_V2_lista()
    elif opcao=="IW_SP":
        print(f'Opcao escolhida: {opcao}')
        sp_fech_esto_V2_lista()
    else:
        print(f'Opcao Invalida!!!')
        
def Seleciona_query_Detalhe(opcao , id_fech):
    print(f'========= Seleciona_query_Detalhe\n')
    
    if opcao=="IW_ES":
        print(f'Opcao escolhida: {opcao}\nId preenchido:{id_fech}')
        es_fech_esto_V2_detalhado(id_fech)
    elif opcao=="IW_RJ":
        print(f'Opcao escolhida: {opcao}\nId preenchido:{id_fech}')
        rj_fech_esto_V2_detalhado(id_fech)
    elif opcao=="IW_SP":
        print(f'Opcao escolhida: {opcao}\nId preenchido:{id_fech}')
        sp_fech_esto_V2_detalhado(id_fech)
    else:
        print(f'Opcao Invalida!!!')
    

def es_fech_esto_V2_lista():
    th_es_fech_esto_V2_lista = threading.Thread(target=Conect_bd.v2_connection_es_lista()).start()
    th_es_fech_esto_V2_lista_XLSX = pd.read_excel('arquivos\IW_PROD_ES_Lista.xlsx')
    #FORMATANDO DATA:
    th_es_fech_esto_V2_lista_XLSX = th_es_fech_esto_V2_lista_XLSX
    th_es_fech_esto_V2_lista_XLSX['DT_INICIO'] = th_es_fech_esto_V2_lista_XLSX['DT_INICIO'].dt.strftime('%d-%m-%Y')
    th_es_fech_esto_V2_lista_XLSX['DT_FIM'] = th_es_fech_esto_V2_lista_XLSX['DT_FIM'].dt.strftime('%d-%m-%Y')
    print(f'\n\n Data frame th_es_fech_esto_V2_lista_XLSX:\n{th_es_fech_esto_V2_lista_XLSX.info()}')
    label_file_lista.config(text=th_es_fech_esto_V2_lista_XLSX.head(5).to_string(index=False))
    
def rj_fech_esto_V2_lista():
    th_rj_fech_esto_V2_lista = threading.Thread(target=Conect_bd.v2_connection_rj_lista()).start()
    th_rj_fech_esto_V2_lista_XLSX = pd.read_excel('arquivos\IW_PROD_RJ_Lista.xlsx')
    #FORMATANDO DATA:
    th_rj_fech_esto_V2_lista_XLSX = th_rj_fech_esto_V2_lista_XLSX
    th_rj_fech_esto_V2_lista_XLSX['DT_INICIO'] = th_rj_fech_esto_V2_lista_XLSX['DT_INICIO'].dt.strftime('%d-%m-%Y')
    th_rj_fech_esto_V2_lista_XLSX['DT_FIM'] = th_rj_fech_esto_V2_lista_XLSX['DT_FIM'].dt.strftime('%d-%m-%Y')
    print(f'\n\n Data frame th_rj_fech_esto_V2_lista_XLSX:\n{th_rj_fech_esto_V2_lista_XLSX.info()}')
    label_file_lista.config(text=th_rj_fech_esto_V2_lista_XLSX[['ID','DT_INICIO','DT_FIM']].head(5).to_string(index=False))
    
def sp_fech_esto_V2_lista():
    th_sp_fech_esto_V2_lista = threading.Thread(target=Conect_bd.v2_connection_sp_lista()).start()
    th_sp_fech_esto_V2_lista_XLSX = pd.read_excel('arquivos\IW_PROD_SP_Lista.xlsx')
    #FORMATANDO DATA:
    th_sp_fech_esto_V2_lista_XLSX = th_sp_fech_esto_V2_lista_XLSX
    th_sp_fech_esto_V2_lista_XLSX['DT_INICIO'] = th_sp_fech_esto_V2_lista_XLSX['DT_INICIO'].dt.strftime('%d-%m-%Y')
    th_sp_fech_esto_V2_lista_XLSX['DT_FIM'] = th_sp_fech_esto_V2_lista_XLSX['DT_FIM'].dt.strftime('%d-%m-%Y')
    print(f'\n\n Data frame th_sp_fech_esto_V2_lista_XLSX:\n{th_sp_fech_esto_V2_lista_XLSX.info()}')
    label_file_lista.config(text=th_sp_fech_esto_V2_lista_XLSX.head(5).to_string(index=False))

   
#TODO: puxar detalhado do ES:
def es_fech_esto_V2_detalhado(id):   
    print(f"************* es_fech_esto_V2_detalhado -> {id}")
    #retornara um data frame com os dados da query:   
    th_es_fech_esto_V2_detalhado = threading.Thread(target=Conect_bd.v2_connection_es_detalhado(id)).start()
    #ler arquivo xlsx criado e exibir na label:
    th_es_fech_esto_V2_detalhado_XLSX = pd.read_excel('arquivos\IW_PROD_ES_Resultado.xlsx')
    print(th_es_fech_esto_V2_detalhado_XLSX.info())
   
    #th_es_fech_esto_V2_detalhado_XLSX = th_es_fech_esto_V2_detalhado_XLSX[['NOME_MATERIAL' , 'QTDE_SALDO_INIC' , 'VALOR_INIC' , 'VALOR_FINAL' ]]
    #print(f'Data frame apos o XLSX:\n{th_es_fech_esto_V2_detalhado_XLSX.head(5).to_string(index=False)}\n\n')
    
    """
    th_es_fech_esto_V2_detalhado_XLSX = th_es_fech_esto_V2_detalhado_XLSX[['TIPOMATERIAL' , 'VALOR_INIC' , 'VALOR_FINAL' ]]
    print(f'Data frame apos o XLSX:\n{th_es_fech_esto_V2_detalhado_XLSX.head(5).to_string(index=False)}\n\n')
    
    #trazendo os valores distintos das dietas:
    dietas_distintas = th_es_fech_esto_V2_detalhado_XLSX['TIPOMATERIAL'].unique()
    print(f'valores distintos das dietas: {dietas_distintas}')
    
    #Separando as dietas:
    todas_as_dietas = th_es_fech_esto_V2_detalhado_XLSX.query('TIPOMATERIAL == "Dietas"')
    print(f'Sub total dietas inicial: {todas_as_dietas["VALOR_INIC"].sum()}')
    print(f'Sub total dietas final: {todas_as_dietas["VALOR_FINAL"].sum()}')
    
    #separando Mat. Enfermagem:
    todas_as_dietas = th_es_fech_esto_V2_detalhado_XLSX.query('TIPOMATERIAL == "Mat. Enfermagem"')
    print(f'Sub total Mat. Enfermagem inicial: {todas_as_dietas["VALOR_INIC"].sum()}')
    print(f'Sub total Mat. Enfermagem final: {todas_as_dietas["VALOR_FINAL"].sum()}')
    
    #separando Materiais Diversos:
    todas_as_dietas = th_es_fech_esto_V2_detalhado_XLSX.query('TIPOMATERIAL == "Materiais Diversos"')
    print(f'Sub total Materiais Diversos inicial: {todas_as_dietas["VALOR_INIC"].sum()}')
    print(f'Sub total Materiais Diversos final: {todas_as_dietas["VALOR_FINAL"].sum()}')
        
    #separando Medicamentos:
    todas_as_dietas = th_es_fech_esto_V2_detalhado_XLSX.query('TIPOMATERIAL == "Medicamentos"')
    print(f'Sub total Medicamentos inicial: {todas_as_dietas["VALOR_INIC"].sum()}')
    print(f'Sub total Medicamentos final: {todas_as_dietas["VALOR_FINAL"].sum()}')    
    
    print('\n')
    #somando todos os VALOR_INIC das dietas:
    print(f'Total estoque inicial: {th_es_fech_esto_V2_detalhado_XLSX["VALOR_INIC"].sum()}')
    
    #somando todos os VALOR_FINAL das dietas:
    print(f'Total estoque final: {th_es_fech_esto_V2_detalhado_XLSX["VALOR_FINAL"].sum()}')
    
    """
    
    #resultados inseridos num array:
    sub_totais = {
                    'Sub total dietas inicial' : th_es_fech_esto_V2_detalhado_XLSX.query('TIPOMATERIAL == "Dietas"')["VALOR_INIC"].sum(),
                    'Sub total dietas final' : th_es_fech_esto_V2_detalhado_XLSX.query('TIPOMATERIAL == "Dietas"')["VALOR_FINAL"].sum(),
                    'Sub total Mat. Enfermagem inicial' : th_es_fech_esto_V2_detalhado_XLSX.query('TIPOMATERIAL == "Mat. Enfermagem"')["VALOR_INIC"].sum(),
                    'Sub total Mat. Enfermagem final' : th_es_fech_esto_V2_detalhado_XLSX.query('TIPOMATERIAL == "Mat. Enfermagem"')["VALOR_FINAL"].sum(),
                    'Sub total Materiais Diversos inicial' : th_es_fech_esto_V2_detalhado_XLSX.query('TIPOMATERIAL == "Materiais Diversos"')["VALOR_INIC"].sum(),
                    'Sub total Materiais Diversos final' : th_es_fech_esto_V2_detalhado_XLSX.query('TIPOMATERIAL == "Materiais Diversos"')["VALOR_FINAL"].sum(),
                    'Sub total Medicamentos inicial' : th_es_fech_esto_V2_detalhado_XLSX.query('TIPOMATERIAL == "Medicamentos"')["VALOR_INIC"].sum(),
                    'Sub total Medicamentos final' : th_es_fech_esto_V2_detalhado_XLSX.query('TIPOMATERIAL == "Medicamentos"')["VALOR_FINAL"].sum(),
                    'Total estoque inicial' : th_es_fech_esto_V2_detalhado_XLSX["VALOR_INIC"].sum(),
                    'Total estoque final' : th_es_fech_esto_V2_detalhado_XLSX["VALOR_FINAL"].sum()
                }
    print(f'\n\nDicionario com resultados:\n{sub_totais}')
    
    #resultados inseridos em um data frame:
    df_resultados = pd.DataFrame.from_dict(sub_totais, orient='index', columns=['Valor'])
    print(f'\nResultados:\n{df_resultados.info}')   
    
    #exibindo o data frame na tela:
    label_file_dados.config(text=th_es_fech_esto_V2_detalhado_XLSX.head(5).to_string(index=False))
    
    #TODO: refatorar
    label_file_dados.config(text=df_resultados)


    
def rj_fech_esto_V2_detalhado(id):   
    print(f"************* rj_fech_esto_V2_detalhado -> {id}")
    #retornara um data frame com os dados da query:   
    th_rj_fech_esto_V2_detalhado = threading.Thread(target=Conect_bd.v2_connection_rj_detalhado(id)).start()
    #ler arquivo xlsx criado e exibir na label:
    th_rj_fech_esto_V2_detalhado_XLSX = pd.read_excel('arquivos\IW_PROD_RJ_Resultado.xlsx')
    print(th_rj_fech_esto_V2_detalhado_XLSX.info())
    
    #exibindo o data frame na tela:
    th_rj_fech_esto_V2_detalhado_XLSX = th_rj_fech_esto_V2_detalhado_XLSX[['NOME_MATERIAL' , 'QTDE_SALDO_INIC' , 'VALOR_INIC' , 'VALOR_FINAL' ]]
    label_file_dados.config(text=th_rj_fech_esto_V2_detalhado_XLSX.head(5).to_string(index=False))
    
    
def sp_fech_esto_V2_detalhado(id):   
    print(f"************* sp_fech_esto_V2_detalhado -> {id}")
    #retornara um data frame com os dados da query:   
    th_sp_fech_esto_V2_detalhado = threading.Thread(target=Conect_bd.v2_connection_sp_detalhado(id)).start()
    #ler arquivo xlsx criado e exibir na label:
    th_sp_fech_esto_V2_detalhado_XLSX = pd.read_excel('arquivos\IW_PROD_SP_Resultado.xlsx')
    print(th_sp_fech_esto_V2_detalhado_XLSX.info())
    
    #exibindo o data frame na tela:
    th_sp_fech_esto_V2_detalhado_XLSX = th_sp_fech_esto_V2_detalhado_XLSX[['NOME_MATERIAL' , 'QTDE_SALDO_INIC' , 'VALOR_INIC' , 'VALOR_FINAL' ]]
    label_file_dados.config(text=th_sp_fech_esto_V2_detalhado_XLSX.head(5).to_string(index=False))
    
        

#============================================================ EXECUCAO ============================================================
if __name__ == "__main__":
    try:
        print("\n============================== inicio ========================")

        #montar interface gráfica:
        root = tk.Tk()
        #root.geometry("1024x800")
        root.minsize(800,700)
        root.maxsize(1000,900)
        root.title("ROBO - Fechamento Suprimentos 1.0")
        root.configure(bg="white")
       
        imagem_tk = ImageTk.PhotoImage(imagem_redimencionada)
        lb_barra_superior = tk.Label(root, image=imagem_tk , border =0)
        lb_barra_superior.pack(pady=10)
        
        opcao_var = tk.StringVar()
        opcao_var.set("Nenhuma opcao!")
    
        radio_es = tk.Radiobutton(root ,text="IW ES", variable=opcao_var, value='IW_ES' ,command=lambda: [print(f'RadionButton: {opcao_var.get()}'), Unidade:=opcao_var.get(), print(f'Unidade: {Unidade}')])
        radio_es.pack()
        
        radio_rj = tk.Radiobutton(root ,text="IW RJ", variable=opcao_var , value='IW_RJ' ,command=lambda: [print(f'RadionButton: {opcao_var.get()}') , Unidade:=opcao_var.get(), print(f'Unidade: {Unidade}')])
        radio_rj.pack()
    
        radio_sp = tk.Radiobutton(root ,text="IW SP", variable=opcao_var, value='IW_SP' ,command=lambda: [print(f'RadionButton: {opcao_var.get()}'), Unidade:=opcao_var.get(), print(f'Unidade: {Unidade}')])
        radio_sp.pack()
        
        #botao para exibicao de lista de fechamentos:
        bt_list_V2_fech_lista = tk.Button(root, text="Lista Fechamentos V2" , command=lambda: [Unidade:=opcao_var.get() , print(f'botao bt_list_V2_fech_esto: {Unidade}') , Seleciona_query_List(Unidade) ])
        bt_list_V2_fech_lista.pack(pady=10)
        
        label_file_lista = tk.Label(root,text='',border =0)
        label_file_lista.pack(pady=10)

        #botao para V2_fech_esto_RJ
        campo_entrada = tk.Entry(root)
        campo_entrada.pack(pady=10)
        
        
        #TODO: verificar !!!!!!!!!!!!!!!!
        bt_V2_fech_esto_detalh = tk.Button(root, text="Detalhes Fechamento V2" , command=lambda: [ Unidade:=opcao_var.get() , Seleciona_query_Detalhe(Unidade , campo_entrada.get())])
        bt_V2_fech_esto_detalh.pack(pady=10)
        
        
        
        
        #TO DO: exibindo em tela :
        label_file_dados = tk.Label(root,text='',border =0)
        label_file_dados.pack(pady=10)
        
        root.mainloop()
        print("\n============================== fim ========================")
    
    except Exception as err:
        print(f"Erro Inexperado: {err=}, \n{type(err)=}")
        pyautogui.alert(f"Erro Inexperado: {err=}, \n{type(err)=}")