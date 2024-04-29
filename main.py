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


def rj_fech_esto_V2_lista():
    th_rj_fech_esto_V2_lista = threading.Thread(target=Conect_bd.v2_connection_rj_lista()).start()
    th_rj_fech_esto_V2_lista_XLSX = pd.read_excel('arquivos\IW_PROD_RJ_Lista.xlsx')
    label_file_lista.config(text=th_rj_fech_esto_V2_lista_XLSX.head(5).to_string(index=False))
    
    
def rj_fech_esto_V2_detalhado(id):   
    print(f"************* rj_fech_esto_V2_detalhado -> {id}")
    #retornara um data frame com os dados da query:   
    th_rj_fech_esto_V2_detalhado = threading.Thread(target=Conect_bd.v2_connection_rj_detalhado(id)).start()
          
    #ler arquivo xlsx criado e exibir na label:
    th_rj_fech_esto_V2_detalhado_XLSX = pd.read_excel('arquivos\IW_PROD_RJ_Resultado.xlsx')
    label_file_dados.config(text=th_rj_fech_esto_V2_detalhado_XLSX.head(5))
    
        

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
    
        radio_es = tk.Radiobutton(root ,text="IW ES", variable=opcao_var, value='IW_ES' ,command=lambda: [print(f'RadionButton: {opcao_var.get()}')])
        radio_es.pack()
        
        radio_rj = tk.Radiobutton(root ,text="IW RJ", variable=opcao_var, value='IW_RJ' ,command=lambda: [print(f'RadionButton: {opcao_var.get()}') , Unidade:=opcao_var.get(), print(f'Unidade: {Unidade}')])
        radio_rj.pack()
    
        radio_sp = tk.Radiobutton(root ,text="IW SP", variable=opcao_var, value='IW_SP' ,command=lambda: [print(f'RadionButton: {opcao_var.get()}')])
        radio_sp.pack()
        
        #botao para exibicao de lista de fechamentos:
        bt_list_V2_fech_esto_RJ = tk.Button(root, text="Lista Fechamentos V2" , command=lambda: [print('botao bt_list_V2_fech_esto_RJ!!!!') , rj_fech_esto_V2_lista() ])
        bt_list_V2_fech_esto_RJ.pack(pady=10)
        
        label_file_lista = tk.Label(root,text='',border =0)
        label_file_lista.pack(pady=10)

        #botao para V2_fech_esto_RJ
        campo_entrada = tk.Entry(root)
        campo_entrada.pack(pady=10)
        bt_V2_fech_esto_RJ = tk.Button(root, text="Detalhes Fechamento V2" , command=lambda: [rj_fech_esto_V2_detalhado(campo_entrada.get())])
        bt_V2_fech_esto_RJ.pack(pady=10)
        
        #TO DO: exibindo em tela :
        label_file_dados = tk.Label(root,text='',border =0)
        label_file_dados.pack(pady=10)
        
        root.mainloop()
        print("\n============================== fim ========================")
    
    except Exception as err:
        print(f"Erro Inexperado: {err=}, \n{type(err)=}")