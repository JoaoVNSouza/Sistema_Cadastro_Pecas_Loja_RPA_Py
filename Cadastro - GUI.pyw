# Importar bibliotecas necessárias.
import xml.etree.ElementTree as ET
import pandas as pd
import pyautogui as pag
import time as t
import customtkinter
from func import *
import os

# Algumas definições.
pag.PAUSE = 0.3         # Tempo de delay entre os comandos da automação.
pag.FAILSAFE = True     # Encerrar o programa se o mouse no canto "up left".

# Local padrão para os arquivos.
os.chdir(absolute_path)


# Função do botão gerar_excel.
def gerar_excel(window):
    marca = window.combobox_1.get().upper()    # Marca em Maiúsculo.
    porcentagem = window.entry.get()           # Pega o valor de porcentagem.

    if porcentagem:
        porcentagem = int(porcentagem)/100 # Valor da porcentagem.

    # Casos especiais.
    if marca == 'MARCA8':
        marca = 'MARCA1'
        tag = 1
    elif marca == 'MARCA7':
        tag = 2
    else:
        tag = None

    #if True:
    try:
        if marca != 'MARCA3':
            # Criando objeto com o XML.
            tree = ET.parse('nota.xml')
            root = tree.getroot()
            nsNFE = {'ns': "http://www.portalfiscal.inf.br/nfe"}  # Dicionário.

            NCM, desc, cod, ref, qtd, preco, grupo_cod, grupo_nome, modify = extrair_dados(
                root, nsNFE, marca, tag, porcentagem)

        else:
            marca3_df = pd.read_excel('nota.xlsx')  # marca3_df

            NCM, desc, cod, ref, qtd, preco, grupo_cod, grupo_nome, modify = marca3(
                marca3_df)

        # Criando dataFrame.
        produtos = pd.DataFrame(data={'NCM': NCM, 'Grupo nome': grupo_nome, 'Grupo cód': grupo_cod, 'Código prod': cod, 'Referência prod': ref,
                                      'Descricao': desc, 'Marca': marca, 'Preço prod': preco, 'Quantidade': qtd, 'Modificar': modify})
        produtos.index += 1  # Resetar os índices a partir de 1.

        produtos.to_excel('produtos.xlsx')  # Exportar dataFrame.

        # Mostrar aviso.
        pag.alert('Gerou arquivo Produtos!', title='AVISO!')
        window.destroy()
    except:
        pag.alert('Arquivo não encontrado ou alguma falha aparente!',
                  title='AVISO!')


# Função do botão cadastrar.
def pag_cadastrar(window2):
    try:
        # Ler dados.
        try:
            produtos = pd.read_excel(
                'FaltaCadastrar.xlsx', index_col=0, dtype=str)
            # Resetar os índices a partir de 1.
            produtos.reset_index(drop=True, inplace=True)
            produtos.index += 1
            # Exportar df.
            produtos.to_excel('FaltaCadastrar.xlsx')
        except:
            produtos = pd.read_excel('Produtos.xlsx', index_col=0, dtype=str)
            # Elimina linhas vazias.
            produtos.dropna(axis=0, inplace=True, how='all')
            # Resetar os índices a partir de 1.
            produtos.reset_index(drop=True, inplace=True)
            produtos.index += 1
            # Exportar df.
            produtos.to_excel('Produtos.xlsx')
            produtos.to_excel('FaltaCadastrar.xlsx')

        # Abrir o programa que já está aberto na tela inicial.
        t.sleep(1)
        pag.click(x=163, y=739)  # Abre o programa.
        pag.click(x=30, y=30)    # Abre as opções de cadastro.
        pag.click(x=50, y=135)   # Abre o menu cadastrar produtos.
        t.sleep(2)               # Pausa de 2 segundos para começar.

        # Verifica se o caps lock está ativado.
        if caps_lock_on():
            pag.press('capslock')  # Desativa a tecla caps lock.

        # Cadastro peça por peças.
        c = 1
        for _ in range(len(produtos)):
            produtos = pd.read_excel(
                'FaltaCadastrar.xlsx', index_col=0, dtype=str)

            # Pegar dados da tabela.
            NCM = produtos['NCM'][c]
            ref = str(produtos['Referência prod'][c])
            desc = produtos['Descricao'][c].upper()
            marca = produtos['Marca'][c]
            preco = produtos['Preço prod'][c]
            # qtd = produtos['Quantidade'][c]   # Não precisa nessa versão.
            grupo = produtos['Grupo cód'][c]

            # Automação.
            pag.press('tab')
            pag.write(NCM)
            pag.press('tab')
            pag.write(grupo)
            for i in range(5):          # 5 x tab
                pag.press('tab')
            pag.write('n')
            pag.press('tab')
            pag.write(ref)
            pag.press('tab')
            pag.write(desc)
            pag.press('tab')
            pag.write(marca)
            pag.click(x=364, y=545)
            pag.write(preco)
            for i in range(3):          # 3 x tab
                pag.press('tab')
            # pag.write(qtd) # Não inserir quantidade nessa versão.
            pag.press('f4')
            produtos.drop(c, axis=0, inplace=True)
            produtos.to_excel('FaltaCadastrar.xlsx')
            c += 1

        # Depois da última peça.
        pag.press('esc')
        os.remove('FaltaCadastrar.xlsx')
        pag.alert(
            'Cadastro Finalizado! Já pode utilizar o computador!', title='AVISO')
        window2.destroy()
    except:
        pag.alert('Automação cancelada, veja o arquivo \"Falta Cadastrar\"\nCom os produtos que ainda não foram cadastrados!', title='AVISO')


# Função etiquetar peças.
def pag_etiquetar(window3):
    try:
        # Abrir o programa e a opção de etiquetar.
        t.sleep(1)
        pag.click(x=163, y=739)               # Abre o programa.
        pag.click(x=163, y=33)                # Abre o Mov. Diário.
        pag.click(x=229, y=450)               # Abre Etiquetas.
        pag.click(x=427, y=450)               # Abre Produtos

        # Configurar opção por código -> Fazer antes de iniciar a automação.
        pag.press('*')
        pag.press('enter')
        pag.click(x=61, y=87)
        pag.write('Codigo')
        pag.press('esc')

        codigo = int(window3.entryCod.get())    # Código do 1 produto cadastrado.

        produtos = pd.read_excel('Produtos.xlsx', index_col=0, dtype=str)
        # Elimina linhas vazias.
        produtos.dropna(axis=0, inplace=True, how='all')
        # Resetar os índices a partir de 1.
        produtos.reset_index(drop=True, inplace=True)
        produtos.index += 1

        c = 1
        for _ in range(len(produtos)):
            pag.press('*')
            pag.press('enter')
            pag.write(str(codigo))
            pag.press('enter')
            for i in range(4):
                pag.press('enter')
            pag.write(str(produtos['Quantidade'][c]))
            for i in range(2):
                pag.press('enter')

            c += 1          # Incrementa o produto atual.
            codigo += 1     # Incrementa o código.

        # Imprimir etiquetas.
        t.sleep(1)
        pag.click(x=405, y=407, clicks=1)

        # Apagar etiquetas.
        t.sleep(1)
        pag.click(x=508, y=364, clicks=1)
        t.sleep(1)
        pag.press('enter')
        t.sleep(1)
        pag.press('enter')
        t.sleep(1)
        pag.press('esc')

        pag.alert(
            'Etiquetas retirados com sucesso! Já pode utilizar o computador!', title='AVISO')
        window3.destroy()
    except:
        pag.alert(
            '\tAutomação cancelada.\nNem todos as etiquetas fram impressas!', title='AVISO')


# Configurações básicas da janela.
customtkinter.set_appearance_mode('system')
customtkinter.set_default_color_theme("green")


class App(customtkinter.CTk):
    # Tela principal.
    def __init__(self):
        super().__init__()

        # configure window
        self.title("Sistema")
        self.geometry(f"{800}x{250}+{285}+{200}")
        self.resizable(False, False)
        fontStyle = customtkinter.CTkFont(family="Arial", size=20)

        # configure grid layout (4x4)
        self.grid_columnconfigure(1, weight=1)
        self.grid_columnconfigure((2, 3), weight=0)
        self.grid_rowconfigure((0, 1, 2), weight=1)

        self.text = customtkinter.CTkLabel(self, text="")
        self.text.grid(row=0, column=0, padx=25, pady=100)

        self.button = customtkinter.CTkButton(self, text="1º Gerar excel",
                                              command=self.tela_excel, width=160, height=80, font=fontStyle,
                                              text_color='white', corner_radius=20)

        self.button.grid(row=0, column=1, padx=0, pady=0)

        self.button = customtkinter.CTkButton(self, text="2º Cadastrar",
                                              command=self.tela_cadastrar, width=160, height=80, font=fontStyle,
                                              text_color='white', corner_radius=20)

        self.button.grid(row=0, column=2, padx=60, pady=0)

        self.button = customtkinter.CTkButton(self, text="3º Etiquetar",
                                              command=self.tela_etiquetar, width=160, height=80, font=fontStyle,
                                              text_color='white', corner_radius=20)

        self.button.grid(row=0, column=3, padx=30, pady=0)

        self.text = customtkinter.CTkLabel(self, text="")
        self.text.grid(row=0, column=4, padx=25, pady=0)

    # Tela Gerar_excel.
    def tela_excel(self):
        window = customtkinter.CTkToplevel(self)
        fontStyle = customtkinter.CTkFont(family="Arial", size=20)

        # configure window
        window.title("Gerar Excel")
        window.geometry("{}x{}+{}+{}".format(700, 280,
                        self.winfo_x() + 50, self.winfo_y() + 50))
        window.grab_set()
        window.resizable(False, False)

        window.text = customtkinter.CTkLabel(
            window, text="", font=customtkinter.CTkFont(size=15, weight="bold"))
        window.text.grid(row=0, column=0, padx=10, pady=(20, 10))

        window.text1 = customtkinter.CTkLabel(
            window, text="Antes de iniciar, insira o arquivo \"nota.xml\"  ou", font=customtkinter.CTkFont(size=15, weight="bold"))
        window.text1.grid(row=0, column=1, padx=0, pady=(20, 10))

        window.text1 = customtkinter.CTkLabel(
            window, text="  \"nota.xlsx\" na pasta \"Arquivos\"", font=customtkinter.CTkFont(size=15, weight="bold"))
        window.text1.grid(row=0, column=2, padx=0, pady=(20, 10))

        window.text2 = customtkinter.CTkLabel(
            window, text="Escolha a marca:", font=customtkinter.CTkFont(size=15, weight="bold"))
        window.text2.grid(row=1, column=1, padx=20, pady=(20, 10))

        window.combobox_1 = customtkinter.CTkComboBox(
            window, values=["Marca1", "Marca2", "Marca3", "Marca4", "Marca5", "Marca6", 'Marca7', 'Marca8'])
        window.combobox_1.grid(row=1, column=2, padx=0, pady=(20, 10))

        window.text = customtkinter.CTkLabel(
            window, text="", font=customtkinter.CTkFont(size=15, weight="bold"))
        window.text.grid(row=2, column=3, padx=20, pady=(20, 10))

        window.text3 = customtkinter.CTkLabel(
            window, text="Degite a porcentagem:", font=customtkinter.CTkFont(size=15, weight="bold"))
        window.text3.grid(row=2, column=1, padx=20, pady=(20, 10))

        window.entry = customtkinter.CTkEntry(window)
        window.entry.grid(row=2, column=2, padx=20, pady=(20, 10))

        window.text = customtkinter.CTkLabel(
            window, text="", font=customtkinter.CTkFont(size=15, weight="bold"))
        window.text.grid(row=3, column=3)

        window.confirmar = customtkinter.CTkButton(
            window, text="Confirmar", command=lambda: gerar_excel(window), fg_color='green', font=fontStyle)
        window.confirmar.grid(row=4, column=1, padx=20, pady=(10, 10))

        window.cancelar = customtkinter.CTkButton(
            window, text="Cancelar", command=lambda: window.destroy(), fg_color='red', font=fontStyle)
        window.cancelar.grid(row=4, column=2, padx=20, pady=(10, 10))

        window.text = customtkinter.CTkLabel(
            window, text="", font=customtkinter.CTkFont(size=15, weight="bold"))
        window.text.grid(row=5, column=0, padx=20, pady=(20, 10))

    # Tela cadastrar.
    def tela_cadastrar(self):
        window2 = customtkinter.CTkToplevel(self)
        fontStyle = customtkinter.CTkFont(family="Arial", size=20)

        # configure window
        window2.title("Cadastrar automático")
        window2.geometry("{}x{}+{}+{}".format(700, 280,
                                              self.winfo_x() + 50, self.winfo_y() + 50))
        window2.grab_set()
        window2.resizable(False, False)

        window2.text = customtkinter.CTkLabel(
            window2, text="Deseja iniciar o cadastro automático?\n OBS: Deixe o programa da loja aberto na página inicial.", font=customtkinter.CTkFont(size=15, weight="bold"))
        window2.text.grid(row=0, column=0, padx=120, pady=30)

        window2.Continuar = customtkinter.CTkButton(
            window2, text="Continuar", command=lambda: pag_cadastrar(window2), fg_color='green', font=fontStyle)
        window2.Continuar.grid(row=2, column=0, padx=0, pady=(20, 10))

        window2.cancelar = customtkinter.CTkButton(
            window2, text="Cancelar", command=lambda: window2.destroy(), fg_color='red', font=fontStyle)
        window2.cancelar.grid(row=3, column=0, padx=0, pady=(20, 10))

    # Tela cadastrar.
    def tela_etiquetar(self):
        window3 = customtkinter.CTkToplevel(self)
        fontStyle = customtkinter.CTkFont(family="Arial", size=20)

        # configure window
        window3.title("Etiquetar peças")
        window3.geometry("{}x{}+{}+{}".format(680, 350,
                                              self.winfo_x() + 50, self.winfo_y() + 50))
        window3.grab_set()
        window3.resizable(False, False)

        window3.text = customtkinter.CTkLabel(
            window3, text="Deseja iniciar a etiquetagem automática?\n OBS: Deixe o programa da loja aberto na página inicial.", font=customtkinter.CTkFont(size=15, weight="bold"))
        window3.text.grid(row=0, column=0, padx=120, pady=30)

        window3.text = customtkinter.CTkLabel(
            window3, text=" Digite o código do primeiro produto cadastrado: ", font=customtkinter.CTkFont(size=15, weight="bold"), fg_color='grey', text_color='black')
        window3.text.grid(row=1, column=0)

        window3.entryCod = customtkinter.CTkEntry(
            window3, placeholder_text="Código", border_width=2, corner_radius=10, width=150, height=40)
        window3.entryCod.grid(row=2, column=0, pady=5)

        window3.Continuar = customtkinter.CTkButton(
            window3, text="Continuar", command=lambda: pag_etiquetar(window3), fg_color='green', font=fontStyle)
        window3.Continuar.grid(row=3, column=0, padx=0, pady=(30, 10))

        window3.cancelar = customtkinter.CTkButton(
            window3, text="Cancelar", command=lambda: window3.destroy(), fg_color='red', font=fontStyle)
        window3.cancelar.grid(row=4, column=0, padx=0, pady=(20, 10))


if __name__ == "__main__":
    app = App()
    app.mainloop()
