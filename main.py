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
    label_file_path.config(text=filename , width=100 , height=1)
    label_file_path.pack(pady=10)
    
    #B) Abrir planilha xlsx e exibir em tela:
    
    #TO DO: refatorar linha abaixo, pois ela se trata exclusivamente de teste:
    #data_frame = pd.read_excel('C:/Pietro/Projetos/ROBO-Fechamento_Suprimentos/arquivos/Fechamento Estoque v2 - para export 01012024 ate 31012024 000.xlsx',index_col=None)
    
    #TO DO: linha abaixo importa o data frame corretamente:
    data_frame = pd.read_excel(filename)
    print(f'#Default:\nData Frame:\n{data_frame.head(5)}\n')
    
    
    #mudando rotulos Unnamed:
    data_frame.columns = data_frame.iloc[0]
    
    #removendo coluna 0 com valor nan
    data_frame = data_frame.iloc[:,1:]
    #removendo linha 0 com valor nan
    data_frame = data_frame.iloc[:,1:]
    print(f'#mudando rotulos e removendo NAN:\nData Frame:\n{data_frame}\n')
    
    #TO DO alterando tipos object para número:
    #Converter para float:
    """
    data_frame['CD'] = pd.to_numeric(data_frame['CD'], errors='coerce')
    data_frame['COD_MATERIAL'] = pd.to_numeric(data_frame['COD_MATERIAL'], errors='coerce')
    data_frame['QTDE_SALDO_INIC'] = pd.to_numeric(data_frame['QTDE_SALDO_INIC'], errors='coerce')
    data_frame['QTDE_SALDO_FINAL'] = pd.to_numeric(data_frame['QTDE_SALDO_FINAL'], errors='coerce')
    
    data_frame['QTDE_E_TOTAL'] = pd.to_numeric(data_frame['QTDE_E_TOTAL'], errors='coerce')
    data_frame['QTDE_S_TOTAL'] = pd.to_numeric(data_frame['QTDE_S_TOTAL'], errors='coerce')
    data_frame['QTDE_E_NF'] = pd.to_numeric(data_frame['QTDE_E_NF'], errors='coerce')
    data_frame['QTDE_S_PAC'] = pd.to_numeric(data_frame['QTDE_S_PAC'], errors='coerce')
    data_frame['QTDE_E_INI_SD'] = pd.to_numeric(data_frame['QTDE_E_INI_SD'], errors='coerce')
    """
    
    
    #TO DO alterando tipos object para DATA:
    #data_frame['DATAT0'] = pd.to_datetime(data_frame['DATAT0'])
    
    
    #EXIBINDO EN CONSOLE:    
    #exibindo o index do data frame:
    #print(f'Data frame index():\n{data_frame.index}\n')
    
    #exibindo informacao do data_frame:
    #print(f'\nData Frame info():\n{data_frame.info()}\n\n')
    
    #exibindo a descricao do data frame:
    #print(f'Data Frame describe():\n{data_frame.describe()}\n')
    
    
    #apagando a primeira linha abaixo do titulo:
    data_frame = data_frame.dropna(axis=0, how='all')
    
    #Exibindo colunas específicas:
    selected_columns = ['COD_MATERIAL','NOME_MATERIAL','QTDE_SALDO_INIC']
    #print(f'\nColunas Selecionadas:\n{data_frame[selected_columns]}')
    
    
    #criando dataframe para processamento
    data_pre = data_frame.copy()
    
    #exibindo informacao do data_frame:
    print(f'============== ANTES:\n\nData Frame info():\n{data_pre.info()}\n\n')
    
    data_pre['QTDE_SALDO_INIC'] = pd.to_numeric(data_pre['QTDE_SALDO_INIC'], errors='coerce')
    
    #exibindo informacao do data_frame:
    print(f'============== DEPOIS:\n\nData Frame info():\n{data_pre.info()}\n\n')
    
    
    #exibindo o teste:
    teste = data_pre['QTDE_SALDO_INIC']
    print('TESTE:\n')
    print(teste)
    
    df_temp = data_pre['QTDE_SALDO_INIC']
    print('\n\n\n\n======================ESTATISTICAS')
    print(f'COUNT: {df_temp.count()}')
    print(f'MAX: {df_temp.max()}')
    
    
    
    
    
    #exibicao do data frame na tela:
    #label_file_dados.config(text=data_frame[[ 'COD_MATERIAL', 'NOME_MATERIAL','VALOR_FINAL', 'VALOR_E_TOTAL' , 'VALOR_S_TOTAL']].iloc[1:99999999])
    label_file_dados.config(text=data_pre[['NOME_MATERIAL','QTDE_SALDO_INIC']])
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

        #exibindo em tela o endereco do arquivo selecionado no botao chosefile:
        label_file_path = tk.Label(root,text='',border =0 , width=100 , height=1)
        label_file_path.pack(pady=10)
        
        #botão de chose file:
        bt_choseFile = tk.Button(root, text="Arquivo" , command=UploadAction)
        bt_choseFile.pack(pady=10 ) 
        
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