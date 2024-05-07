#07/05/2024
#@PLima
#setup para criacao de exe do Fechamento Suprimentos

import sys
from cx_Freeze import setup, Executable

#Dependencies are automatically detected, but it might need fine tuning
build_exe_options = {"packages": ["os"], "includes": ["tkinter","PIL","tkinter","pandas","numpy","warnings","pandastable","conect_BD","pyautogui","customtkinter","oracledb","datetime" , "cryptography.hazmat.primitives.kdf.pbkdf2" , "os"], 'include_files': ["LOGO_PRETA.png"]   }



# GUI applications require a different base on Windows (the default is for  
#a console application).
base = None
if sys.platform == "win32":
    base = "Win32GUI"
    
setup(
    name = "Fechamento Suprimentos",
    version = "1.0",
    description = "App para extração de fechamento conforme seleção de ID e Unidade Pronep.",
    options = {"build_exe": build_exe_options},
    executables=[Executable("main.py", base=base, icon ="icone.ico")]
)