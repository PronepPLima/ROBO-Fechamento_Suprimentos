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
import threading
import pyautogui
import customtkinter as ctk

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
    
    #exibindo data frame no label:
    label_file_lista.config(text=th_es_fech_esto_V2_lista_XLSX[['ID','DT_INICIO']].head(5).to_string(index=False) , justify='center', width=20, height=6)
    
def rj_fech_esto_V2_lista():
    th_rj_fech_esto_V2_lista = threading.Thread(target=Conect_bd.v2_connection_rj_lista()).start()
    th_rj_fech_esto_V2_lista_XLSX = pd.read_excel('arquivos\IW_PROD_RJ_Lista.xlsx')
    #FORMATANDO DATA:
    th_rj_fech_esto_V2_lista_XLSX = th_rj_fech_esto_V2_lista_XLSX
    th_rj_fech_esto_V2_lista_XLSX['DT_INICIO'] = th_rj_fech_esto_V2_lista_XLSX['DT_INICIO'].dt.strftime('%d-%m-%Y')
    th_rj_fech_esto_V2_lista_XLSX['DT_FIM'] = th_rj_fech_esto_V2_lista_XLSX['DT_FIM'].dt.strftime('%d-%m-%Y')
    print(f'\n\n Data frame th_rj_fech_esto_V2_lista_XLSX:\n{th_rj_fech_esto_V2_lista_XLSX.info()}')
    label_file_lista.config(text=th_rj_fech_esto_V2_lista_XLSX[['ID','DT_INICIO']].head(5).to_string(index=False) , justify='center', width=20, height=6)
    
def sp_fech_esto_V2_lista():
    th_sp_fech_esto_V2_lista = threading.Thread(target=Conect_bd.v2_connection_sp_lista()).start()
    th_sp_fech_esto_V2_lista_XLSX = pd.read_excel('arquivos\IW_PROD_SP_Lista.xlsx')
    #FORMATANDO DATA:
    th_sp_fech_esto_V2_lista_XLSX = th_sp_fech_esto_V2_lista_XLSX
    th_sp_fech_esto_V2_lista_XLSX['DT_INICIO'] = th_sp_fech_esto_V2_lista_XLSX['DT_INICIO'].dt.strftime('%d-%m-%Y')
    th_sp_fech_esto_V2_lista_XLSX['DT_FIM'] = th_sp_fech_esto_V2_lista_XLSX['DT_FIM'].dt.strftime('%d-%m-%Y')
    print(f'\n\n Data frame th_sp_fech_esto_V2_lista_XLSX:\n{th_sp_fech_esto_V2_lista_XLSX.info()}')
    label_file_lista.config(text=th_sp_fech_esto_V2_lista_XLSX[['ID','DT_INICIO']].head(5).to_string(index=False) , justify='center', width=20, height=6)

   
#TODO: puxar detalhado do ES:
def es_fech_esto_V2_detalhado(id):   
    print(f"************* es_fech_esto_V2_detalhado -> {id}")
    #retornara um data frame com os dados da query:   
    th_es_fech_esto_V2_detalhado = threading.Thread(target=Conect_bd.v2_connection_es_detalhado(id)).start()
    
    #TODO: refatorar todo o codigo abaixo, pois deveria estar dentro de um classe que manipule o data frame e nao na main que é só execucao:
    #ler arquivo xlsx criado e exibir na label:
    th_es_fech_esto_V2_detalhado_XLSX = pd.read_excel('arquivos\IW_PROD_ES_Resultado.xlsx')
    
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

    #Extraindo Chaves e Valores do Dicionário:
    for chave, valor in sub_totais.items():
        descricoes.append(chave)
        valores.append(valor)
    #Criando o DataFrame:
    df_subtotais = pd.DataFrame({"Descrição": descricoes,"Valor": valores})
    # Nomeando as colunas
    df_subtotais.columns = ["Descrição", "Valor"]  
    #Gravando resultados no excel:
    with pd.ExcelWriter('arquivos\IW_PROD_ES_Resultado.xlsx', engine='openpyxl', mode='a') as writer:
        df_subtotais.to_excel(writer, sheet_name='Sintetico', index=False, header=None)
    #exibindo o data frame na tela:
    df_string = df_subtotais.reset_index(drop=True).to_string(index=False, col_space=10)
    label_file_dados.config(text=df_string, justify='right', width=100)
    
    
def rj_fech_esto_V2_detalhado(id):        
    print(f"************* rj_fech_esto_V2_detalhado -> {id}")
    #retornara um data frame com os dados da query:   
    th_rj_fech_esto_V2_detalhado = threading.Thread(target=Conect_bd.v2_connection_rj_detalhado(id)).start()
    
    #TODO: refatorar todo o codigo abaixo, pois deveria estar dentro de um classe que manipule o data frame e nao na main que é só execucao:
    #ler arquivo xlsx criado e exibir na label:
    th_rj_fech_esto_V2_detalhado_XLSX = pd.read_excel('arquivos\IW_PROD_RJ_Resultado.xlsx')
    
    #resultados inseridos num array:
    sub_totais = {
                    'Sub total dietas inicial' : th_rj_fech_esto_V2_detalhado_XLSX.query('TIPOMATERIAL == "Dietas"')["VALOR_INIC"].sum(),
                    'Sub total dietas final' : th_rj_fech_esto_V2_detalhado_XLSX.query('TIPOMATERIAL == "Dietas"')["VALOR_FINAL"].sum(),
                    'Sub total Mat. Enfermagem inicial' : th_rj_fech_esto_V2_detalhado_XLSX.query('TIPOMATERIAL == "Mat. Enfermagem"')["VALOR_INIC"].sum(),
                    'Sub total Mat. Enfermagem final' : th_rj_fech_esto_V2_detalhado_XLSX.query('TIPOMATERIAL == "Mat. Enfermagem"')["VALOR_FINAL"].sum(),
                    'Sub total Materiais Diversos inicial' : th_rj_fech_esto_V2_detalhado_XLSX.query('TIPOMATERIAL == "Materiais Diversos"')["VALOR_INIC"].sum(),
                    'Sub total Materiais Diversos final' : th_rj_fech_esto_V2_detalhado_XLSX.query('TIPOMATERIAL == "Materiais Diversos"')["VALOR_FINAL"].sum(),
                    'Sub total Medicamentos inicial' : th_rj_fech_esto_V2_detalhado_XLSX.query('TIPOMATERIAL == "Medicamentos"')["VALOR_INIC"].sum(),
                    'Sub total Medicamentos final' : th_rj_fech_esto_V2_detalhado_XLSX.query('TIPOMATERIAL == "Medicamentos"')["VALOR_FINAL"].sum(),
                    'Total estoque inicial' : th_rj_fech_esto_V2_detalhado_XLSX["VALOR_INIC"].sum(),
                    'Total estoque final' : th_rj_fech_esto_V2_detalhado_XLSX["VALOR_FINAL"].sum()
                }

    #Extraindo Chaves e Valores do Dicionário:
    for chave, valor in sub_totais.items():
        descricoes.append(chave)
        valores.append(valor)
    #Criando o DataFrame:
    df_subtotais = pd.DataFrame({"Descrição": descricoes,"Valor": valores})
    # Nomeando as colunas
    df_subtotais.columns = ["Descrição", "Valor"]  
    #Gravando resultados no excel:
    with pd.ExcelWriter('arquivos\IW_PROD_RJ_Resultado.xlsx', engine='openpyxl', mode='a') as writer:
        df_subtotais.to_excel(writer, sheet_name='Sintetico', index=False, header=None)
    #exibindo o data frame na tela:
    df_string = df_subtotais.reset_index(drop=True).to_string(index=False, col_space=10)
    label_file_dados.config(text=df_string, justify='right', width=100)
    
def sp_fech_esto_V2_detalhado(id):   
    print(f"************* sp_fech_esto_V2_detalhado -> {id}")
    #retornara um data frame com os dados da query:   
    th_sp_fech_esto_V2_detalhado = threading.Thread(target=Conect_bd.v2_connection_sp_detalhado(id)).start()
    
    #TODO: refatorar todo o codigo abaixo, pois deveria estar dentro de um classe que manipule o data frame e nao na main que é só execucao:
    #ler arquivo xlsx criado e exibir na label:
    th_sp_fech_esto_V2_detalhado_XLSX = pd.read_excel('arquivos\IW_PROD_SP_Resultado.xlsx')
    
    #resultados inseridos num array:
    sub_totais = {
                    'Sub total dietas inicial' : th_sp_fech_esto_V2_detalhado_XLSX.query('TIPOMATERIAL == "Dietas"')["VALOR_INIC"].sum(),
                    'Sub total dietas final' : th_sp_fech_esto_V2_detalhado_XLSX.query('TIPOMATERIAL == "Dietas"')["VALOR_FINAL"].sum(),
                    'Sub total Mat. Enfermagem inicial' : th_sp_fech_esto_V2_detalhado_XLSX.query('TIPOMATERIAL == "Mat. Enfermagem"')["VALOR_INIC"].sum(),
                    'Sub total Mat. Enfermagem final' : th_sp_fech_esto_V2_detalhado_XLSX.query('TIPOMATERIAL == "Mat. Enfermagem"')["VALOR_FINAL"].sum(),
                    'Sub total Materiais Diversos inicial' : th_sp_fech_esto_V2_detalhado_XLSX.query('TIPOMATERIAL == "Materiais Diversos"')["VALOR_INIC"].sum(),
                    'Sub total Materiais Diversos final' : th_sp_fech_esto_V2_detalhado_XLSX.query('TIPOMATERIAL == "Materiais Diversos"')["VALOR_FINAL"].sum(),
                    'Sub total Medicamentos inicial' : th_sp_fech_esto_V2_detalhado_XLSX.query('TIPOMATERIAL == "Medicamentos"')["VALOR_INIC"].sum(),
                    'Sub total Medicamentos final' : th_sp_fech_esto_V2_detalhado_XLSX.query('TIPOMATERIAL == "Medicamentos"')["VALOR_FINAL"].sum(),
                    'Total estoque inicial' : th_sp_fech_esto_V2_detalhado_XLSX["VALOR_INIC"].sum(),
                    'Total estoque final' : th_sp_fech_esto_V2_detalhado_XLSX["VALOR_FINAL"].sum()
                }

    #Extraindo Chaves e Valores do Dicionário:
    for chave, valor in sub_totais.items():
        descricoes.append(chave)
        valores.append(valor)
    #Criando o DataFrame:
    df_subtotais = pd.DataFrame({"Descrição": descricoes,"Valor": valores})
    # Nomeando as colunas
    df_subtotais.columns = ["Descrição", "Valor"]  
    #Gravando resultados no excel:
    with pd.ExcelWriter('arquivos\IW_PROD_SP_Resultado.xlsx', engine='openpyxl', mode='a') as writer:
        df_subtotais.to_excel(writer, sheet_name='Sintetico', index=False, header=None)
    #exibindo o data frame na tela:
    df_string = df_subtotais.reset_index(drop=True).to_string(index=False, col_space=10)
    label_file_dados.config(text=df_string, justify='right', width=100)
    
        



#============================================================ EXECUCAO ============================================================
if __name__ == "__main__":
    
    try:

        print("\n============================== inicio ========================")

        #montar interface gráfica:
        root = tk.Tk()
        root.geometry("800x700")
        root.minsize(800,700)
        root.maxsize(800,700)
        #root.maxsize(1000,900)
        root.title("ROBO - Fechamento Suprimentos 1.0")
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
    
        radio_es = tk.Radiobutton(root ,text="IW ES", variable=opcao_var, value='IW_ES' ,command=lambda: [print(f'RadionButton: {opcao_var.get()}'), Unidade:=opcao_var.get(), print(f'Unidade: {Unidade}')])
        radio_es.place(x=195 , y=190)
        
        radio_rj = tk.Radiobutton(root ,text="IW RJ", variable=opcao_var , value='IW_RJ' ,command=lambda: [print(f'RadionButton: {opcao_var.get()}') , Unidade:=opcao_var.get(), print(f'Unidade: {Unidade}')])
        radio_rj.place(x=195 , y=210)
    
        radio_sp = tk.Radiobutton(root ,text="IW SP", variable=opcao_var, value='IW_SP' ,command=lambda: [print(f'RadionButton: {opcao_var.get()}'), Unidade:=opcao_var.get(), print(f'Unidade: {Unidade}')])
        radio_sp.place(x=195 , y=230)
        
        #botao para exibicao de lista de fechamentos:
        bt_list_V2_fech_lista = tk.Button(root, text="Lista Fechamentos V2" , command=lambda: [Unidade:=opcao_var.get() , print(f'botao bt_list_V2_fech_esto: {Unidade}') , Seleciona_query_List(Unidade) ])
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
        label_campo_entrada = tk.Label(root, text="Digite o ID:").place(x=10 , y=310)
        campo_entrada = tk.Entry(root, width=3)
        campo_entrada.place(x=80 , y=310)
        
        bt_V2_fech_esto_detalh = tk.Button(root, text="Buscar Fechamento V2" , command=lambda: [ Unidade:=opcao_var.get() , Seleciona_query_Detalhe(Unidade , campo_entrada.get())])
        bt_V2_fech_esto_detalh.place(x=15,y=340)
        
        label_file_dados = tk.Label(root,text='',border =0)
        label_file_dados.place(x=150 , y=310)
        
        root.mainloop()
        print("\n============================== fim ========================")

    
    
        """
        print("\n============================== inicio ========================")

        #montar interface gráfica:
        root = ctk.CTk()
        root.geometry("1024x800")
        #root.minsize(800,700)
        #root.maxsize(1000,900)
        root.title("ROBO - Fechamento Suprimentos 1.0")
        #root.configure(bg="white")
        
        
        frame1 = ctk.CTkFrame(master=root , width=1020 , height=250 , fg_color="red").place(x=2,y=5)
        frame2 = ctk.CTkFrame(root , width=1020 , height=250 , fg_color="green").place(x=2,y=260)
        frame3 = ctk.CTkFrame(root , width=1020 , height=250 , fg_color="blue").place(x=2,y=520)
        root.mainloop() 
        """
        
          
    
    except Exception as err:
        print(f"Erro Inexperado: {err=}, \n{type(err)=}")
        pyautogui.alert(f"Erro Inexperado: {err=}, \n{type(err)=}")