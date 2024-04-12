#10/04/2023
#@PLima


import tkinter as tk
from PIL import Image, ImageTk
from tkinter import filedialog

import pandas as pd
import numpy as py

import warnings


from pandastable import Table

#ignorando alertas exibidos:
warnings.filterwarnings("ignore")

#carregando imagem original:
imagem_original = Image.open("LOGO_PRETA.png")
largura=300
altura=140
#redimencionando a imagem:
imagem_redimencionada = imagem_original.resize((largura,altura))

    

def UploadAction(event=None):
    filename = filedialog.askopenfilename()
    print(f"Endereco do arquivo: {filename}\n")
    
    #apos o click, exibicao do endereco do arquivo em console:
    label_file_path.config(text=filename)
    label_file_path.pack(pady=10)
    
    #B) Abrir planilha xlsx e exibir em tela:
    
    #TO DO: refatorar linha abaixo, pois ela se trata exclusivamente de teste:
    data_frame = pd.read_excel('C:/Pietro/Projetos/ROBO-Fechamento_Suprimentos/arquivos/Fechamento Estoque v2 - para export 01012024 ate 31012024 000.xlsx',index_col=None)
    
    #TO DO: linha abaixo importa o data frame corretamente:
    #data_frame = pd.read_excel(filename)
    
    data_frame.columns = data_frame.iloc[0]
    print(data_frame[[ 'COD_MATERIAL', 'NOME_MATERIAL','VALOR_FINAL']].iloc[1:99999999])

    label_file_dados.config(text=data_frame[[ 'COD_MATERIAL', 'NOME_MATERIAL','VALOR_FINAL', 'VALOR_E_TOTAL' , 'VALOR_S_TOTAL']].iloc[1:99999999])
    label_file_dados.pack(pady=10)
    

#============================================================ EXECUCAO ============================================================
if __name__ == "__main__":
    try:
        print("\n============================== inicio ========================")

        #montar interface gráfica:
        root = tk.Tk()
        #root.maxsize(800,600)
        root.geometry("1024x800")
        root.maxsize(1024,1024)
        root.title("ROBO - Fechamento Suprimentos 1.0")
        root.configure(bg="white")

        imagem_tk = ImageTk.PhotoImage(imagem_redimencionada)
        lb_barra_superior = tk.Label(root, image=imagem_tk , border =0)
        lb_barra_superior.pack()

        #botão de chose file:
        bt_choseFile = tk.Button(root, text="Escolher arquivo" , command=UploadAction)
        bt_choseFile.pack(pady=10) 

        #exibindo em tela o endereco do arquivo selecionado no botao chosefile:
        label_file_path = tk.Label(root,text='',border =0)
        
        #exibindo o número de páginas do pdf:
        #label_file_num_pag = tk.Label(root,text='',border =0)
        

        #TO DO: exibindo em tela pdf com fechamento do mes selecionado:
        label_file_dados = tk.Label(root,text='',border =0)
        
        """
        text_widget = tk.Text(root)
        
        frame = tk.Frame(root)
        frame.pack()
        """

        root.mainloop()
        print("\n============================== fim ========================")
    
    except Exception as err:
        print(f"Erro Inexperado: {err=}, \n{type(err)=}")