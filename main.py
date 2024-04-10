#10/04/2023
#@PLima


import tkinter as tk
from PIL import Image, ImageTk

#carregando imagem original:
imagem_original = Image.open("LOGO_PRETA.png")
largura=300
altura=140
#redimencionando a imagem:
imagem_redimencionada = imagem_original.resize((largura,altura))

if __name__ == "__main__":    
    print("\n============================== inicio ========================")
    
    #montar interface gráfica:
    root = tk.Tk()
    root.maxsize(800,600)
    root.geometry("800x600")
    root.title("ROBO - Fechamento Suprimentos 1.0")
    root.configure(bg="white")
    
    #inserir imagem de logo pronep mas redimencionada
    #imagem = tk.PhotoImage(file="LOGO_PRETA.png")
    #lb_barra_superior = tk.Label(root, image=imagem , border =0 , height=148 , width=320)
    #lb_barra_superior.pack()
    
    imagem_tk = ImageTk.PhotoImage(imagem_redimencionada)
    lb_barra_superior = tk.Label(root, image=imagem_tk , border =0)
    lb_barra_superior.pack()
    
    root.mainloop()
    
    #abrir pasta arquivos, que fica na pasta da instalacao do app
    
    #to do: criar interface gráfica com botão de chose file
    
    #abrir planilha csv com fechamento do mes selecionado
    
    
    
    print("\n============================== fim ========================")