# Bibliotecas utilizadas.
import xml.etree.ElementTree as ET
import pandas as pd
import pyautogui as pag
import pyperclip as clip
import numpy as np
import time as t
import customtkinter as ctk
import os
import ctypes
from typing import List, Tuple

# Constantes.
current_dir = os.getcwd()                  # End. do sistema.
parent_dir = os.path.dirname(current_dir)  # Diretório pai.
path_arquivos = parent_dir + '\Arquivos'   # End. arquivos.
path_imagens = parent_dir + '\Imagens'     # End. imagens.
pag.PAUSE = 0.25                           # Delay entre comandos.
pag.FAILSAFE = True                        # Exit program if mouse "up left".

# Configurações básicas da janela.
ctk.set_appearance_mode('light')
ctk.set_default_color_theme(f"{parent_dir}\\custom_tema.json")
COR_BTN_OK = '#34ac36'
COR_BTN_CANCEL = '#ce0400'

# Local padrão para os arquivos.
os.chdir(path_arquivos)                   # Pasta padrão dos arquivos.
