#10/04/2023
#@PLima


import tkinter as tk
from PIL import Image, ImageTk
from tkinter import filedialog
import pdfplumber

#from PyPDF2 import PdfReader


#carregando imagem original:
imagem_original = Image.open("LOGO_PRETA.png")
largura=300
altura=140
#redimencionando a imagem:
imagem_redimencionada = imagem_original.resize((largura,altura))

def UploadAction(event=None):
    filename = filedialog.askopenfilename()
    print(f"Endereco do arquivo: {filename}")
    #apos o click, exibicao do endereco do arquivo em console:
    label_file_path.config(text=filename)
    
    #abrir pdf com fechamento do mes selecionado  
    with pdfplumber.open(filename) as pdf:
        print(f"\nPaginas: {pdf.pages}\n")
        num_paginas = len(pdf.pages)
        print(f"\nNumero de paginas: {num_paginas}")
        
        #TO DO: ler cada página e exibir no console:       
        primeira_pagina = pdf.pages[0]
        texto = primeira_pagina.extract_text()
    #print(f'Texto: {texto}')
    
    #exibir pdf com fechamento do mes selecionado   
    label_file_dados.config(text=f'{texto}')


#============================================================ EXECUCAO ============================================================
if __name__ == "__main__":
    try:
        print("\n============================== inicio ========================")

        #montar interface gráfica:
        root = tk.Tk()
        root.maxsize(800,600)
        root.geometry("800x600")
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
        label_file_path.pack(pady=10)

        #TO DO: exibindo em tela pdf com fechamento do mes selecionado:
        label_file_dados = tk.Label(root,text='',border =0)
        label_file_dados.pack(pady=10)

        root.mainloop()
        print("\n============================== fim ========================")
    
    except Exception as err:
        print(f"Erro Inexperado: {err=}, \n{type(err)=}")