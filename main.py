#10/04/2023
#@PLima


import tkinter as tk
from PIL import Image, ImageTk
from tkinter import filedialog
import pandas as pd
import numpy as py
import warnings
from pandastable import Table
#from conect_BD import v2_connection_rj
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
    
    
def rj_fech_esto_V2(id):   
    #retornara um data frame com os dados da query:   
    th_rj_fech_esto_V2 = threading.Thread(target=Conect_bd.v2_connection_rj(id)).start()
    
    print(f'THREAD: th_rj_fech_esto_V2\n{th_rj_fech_esto_V2}')
    
    #label_file_dados.config(text=v2_connection_rj(id))
    
    #todo: ler arquivo xlsx criado
    
    #todo: montar um data frame 
    
    #todo: montar esse data frame na label abaixo
    label_file_dados.config(text="Teste")
    label_file_dados.pack(pady=10)
    
        

#============================================================ EXECUCAO ============================================================
if __name__ == "__main__":
    try:
        print("\n============================== inicio ========================")

        #montar interface gráfica:
        root = tk.Tk()
        root.geometry("1024x800")
        root.maxsize(1024,1024)
        root.title("ROBO - Fechamento Suprimentos 1.0")
        root.configure(bg="white")

        imagem_tk = ImageTk.PhotoImage(imagem_redimencionada)
        lb_barra_superior = tk.Label(root, image=imagem_tk , border =0)
        lb_barra_superior.pack()

        #botao para V2_fech_esto_RJ
        id = 44
        
        bt_V2_fech_esto_RJ = tk.Button(root, text="Fechamento V2 - RJ" , command=lambda: [print('Acao do botao RJ!!!!') , rj_fech_esto_V2(id) ])
        bt_V2_fech_esto_RJ.pack(pady=10 ) 
        
        #TO DO: exibindo em tela :
        label_file_dados = tk.Label(root,text='',border =0)
        
        root.mainloop()
        print("\n============================== fim ========================")
    
    except Exception as err:
        print(f"Erro Inexperado: {err=}, \n{type(err)=}")