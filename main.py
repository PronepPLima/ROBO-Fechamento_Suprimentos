#10/04/2023
#@PLima


import tkinter as tk
from PIL import Image, ImageTk
from tkinter import filedialog

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
    display_file_path(filename)   
    
def display_file_path(file_path):
    label_file_path.config(text=file_path)


#============================================================ EXECUCAO ============================================================
if __name__ == "__main__":    
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
    bt_choseFile.pack()
    
    #abrir arquivo selecionado
    
    #abrir planilha csv com fechamento do mes selecionado
    
    label_file_path = tk.Label(root,text='',border =0)
    label_file_path.pack()
    
    root.mainloop()
    

    
    
    
    print("\n============================== fim ========================")