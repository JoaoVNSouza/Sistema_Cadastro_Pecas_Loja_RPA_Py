# Importar bibliotecas necessárias.
from definitions import *
from func_pag import *


class App(ctk.CTk):  # Classe app.
    def __init__(self):  # Tela principal.
        """
        Objetivo: criar a tela principal do programa.

        Parameters:
            self (obj): Objeto da classe.

        Returns:
            None
        """

        super().__init__()

        # Configura a janela.
        self.title("Sistema de cadastro automático")
        self.geometry(f"{800}x{250}+{285}+{200}")
        self.resizable(False, False)
        fontStyle = ctk.CTkFont(family="Arial", size=20)

        # Espaço vazio.
        espaco_vazio = ctk.CTkLabel(self, text="")
        espaco_vazio.grid(row=0, column=0, padx=25, pady=100)

        # Botões.
        btn_excel = ctk.CTkButton(self, text="1º Gerar excel",
                                  command=self.tela_excel, width=160, height=80, font=fontStyle)
        btn_excel.grid(row=0, column=1, padx=0, pady=0)

        btn_cadastrar = ctk.CTkButton(
            self, text="2º Cadastrar", command=self.tela_cadastrar, width=160, height=80, font=fontStyle)
        btn_cadastrar.grid(row=0, column=2, padx=90, pady=0)

        btn_etiquetar = ctk.CTkButton(
            self, text="3º Etiquetar", command=self.tela_etiquetar, width=160, height=80, font=fontStyle)
        btn_etiquetar.grid(row=0, column=3, padx=5, pady=0)

        # Espaço vazio.
        espaco_vazio2 = ctk.CTkLabel(self, text="")
        espaco_vazio2.grid(row=0, column=4, padx=25, pady=0)

        # Botão invísivel para fechar com o ESC.
        btn_cancelar = ctk.CTkButton(self, command=lambda: self.destroy())
        self.bind('<Escape>', lambda event: btn_cancelar.invoke())

    def tela_excel(self):  # Tela Gerar_excel.
        """
        Objetivo: criar a tela de gerar excel.

        Parameters:
            self (obj): Objeto da classe.

        Returns:
            None
        """

        # Configurar janela gerar_excel.
        window_excel = ctk.CTkToplevel(self)
        window_excel.title("Gerar Excel da nota fiscal")
        window_excel.geometry(
            "{}x{}+{}+{}".format(700, 220, self.winfo_x() + 50, self.winfo_y() + 80))
        window_excel.grab_set()
        window_excel.resizable(False, False)

        # Definições.
        fontStyle = ctk.CTkFont(family="Arial", size=20)
        fontStyle2 = ctk.CTkFont(size=15, weight="bold")

        # Espaço vazio.
        espaco_vazio = ctk.CTkLabel(window_excel, text="")
        espaco_vazio.grid(row=0, column=0, padx=12, pady=(20, 10))

        # Texto Aviso.
        text_aviso1 = ctk.CTkLabel(
            window_excel, text="Antes de iniciar, insira o arquivo \"nota.xml\"  ou", font=fontStyle2)
        text_aviso1.grid(row=0, column=1, padx=0, pady=(20, 10))
        text_aviso2 = ctk.CTkLabel(
            window_excel, text="  \"nota.xlsx\" na pasta \"Arquivos\"", font=fontStyle2)
        text_aviso2.grid(row=0, column=2, padx=0, pady=(20, 10))

        # Texto escolher a marca.
        texto_marca = ctk.CTkLabel(
            window_excel, text="Escolha a marca:", font=fontStyle2)
        texto_marca.grid(row=1, column=1, padx=0, pady=(10, 10))

        # Combobox.
        valores = ["Acostamento", "Melissa", "Ilicito",
                   "Index", "Chopper", "Sly", 'Yeskla', 'Act Boné']
        window_excel.combobox = ctk.CTkComboBox(window_excel, values=valores)
        window_excel.combobox.grid(
            row=1, column=2, padx=0, pady=(10, 10), sticky="w")

        # Espaço vazio.
        espaco_vazio2 = ctk.CTkLabel(window_excel, text="")
        espaco_vazio2.grid(row=2, column=3, padx=10, pady=0)

        # Frame para agrupar os botões.
        frame_botoes = ctk.CTkFrame(window_excel)
        frame_botoes.grid(row=3, column=1, columnspan=2)

        # Botões.
        btn_continuar = ctk.CTkButton(frame_botoes, text="Confirmar", command=lambda: gerar_excel(
            window_excel), fg_color=COR_BTN_OK, font=fontStyle, width=100, height=50)
        btn_continuar.pack(side="left", padx=60)
        window_excel.bind('<Return>', lambda event: btn_continuar.invoke())

        btn_cancelar = ctk.CTkButton(frame_botoes, text="Cancelar", command=lambda: window_excel.destroy(
        ), fg_color=COR_BTN_CANCEL, font=fontStyle, width=100, height=50)
        btn_cancelar.pack(side="left", padx=40)
        window_excel.bind('<Escape>', lambda event: btn_cancelar.invoke())

        # Espaço vazio.
        espaco_vazio2 = ctk.CTkLabel(window_excel, text="")
        espaco_vazio2.grid(row=4, column=0, padx=10, pady=(10, 10))

    def tela_cadastrar(self):  # Tela cadastrar.
        """
        Objetivo: criar a tela de cadastrar.

        Parameters:
            self (obj): Objeto da classe.

        Returns:
            None
        """

        # Configurando a janela cadastrar.
        window_cadastrar = ctk.CTkToplevel(self)
        window_cadastrar.title("Cadastrar peças no sistema")
        window_cadastrar.geometry(
            "{}x{}+{}+{}".format(650, 200, self.winfo_x() + 50, self.winfo_y() + 80))
        window_cadastrar.grab_set()
        window_cadastrar.resizable(False, False)

        # Definições.
        fontStyle = ctk.CTkFont(family="Arial", size=20)
        fontStyle2 = ctk.CTkFont(size=15, weight="bold")

        # Texto aviso.
        text_aviso = ctk.CTkLabel(
            window_cadastrar, text="Deseja iniciar o cadastro automático?\n OBS: Deixe o programa da loja aberto na página inicial.", font=fontStyle2)
        text_aviso.grid(row=0, column=0, padx=100, pady=30)

        # Frame para agrupar os botões.
        frame_botoes = ctk.CTkFrame(window_cadastrar)
        frame_botoes.grid(row=1, column=0)

        # Botões.
        btn_continuar = ctk.CTkButton(frame_botoes, text="Continuar", command=lambda: cadastrar(
            window_cadastrar), fg_color=COR_BTN_OK, font=fontStyle, width=100, height=50)
        btn_continuar.pack(side="left", padx=40)
        window_cadastrar.bind(
            '<Return>', lambda event: window_cadastrar.Continuar.invoke())

        btn_cancelar = ctk.CTkButton(frame_botoes, text="Cancelar", command=lambda: window_cadastrar.destroy(
        ), fg_color=COR_BTN_CANCEL, font=fontStyle, width=100, height=50)
        btn_cancelar.pack(side="left", padx=40)
        window_cadastrar.bind('<Escape>', lambda event: btn_cancelar.invoke())

    def tela_etiquetar(self):  # Tela cadastrar.
        """
        Objetivo: criar a tela de cadastrar.

        Parameters:
            self (obj): Objeto da classe.

        Returns:
            None
        """

        # Configurar janela etiquetar.
        window_etiquetar = ctk.CTkToplevel(self)
        window_etiquetar.title("Imprimir etiqueta para as peças cadastradas")
        window_etiquetar.geometry(
            "{}x{}+{}+{}".format(650, 200, self.winfo_x() + 50, self.winfo_y() + 80))
        window_etiquetar.grab_set()
        window_etiquetar.resizable(False, False)

        # Definições.
        fontStyle = ctk.CTkFont(family="Arial", size=20)
        fontStyle2 = ctk.CTkFont(size=15, weight="bold")

        # Texto aviso.
        text_aviso = ctk.CTkLabel(
            window_etiquetar, text="Deseja iniciar a etiquetagem automática?\n OBS: Deixe o programa da loja aberto na página inicial.", font=fontStyle2)
        text_aviso.grid(row=0, column=0, padx=100, pady=30)

        # Frame para agrupar os botões.
        frame_botoes = ctk.CTkFrame(window_etiquetar)
        frame_botoes.grid(row=2, column=0)

        # Botões.
        btn_continuar = ctk.CTkButton(frame_botoes, text="Continuar", command=lambda: cadastrar(
            window_etiquetar), fg_color=COR_BTN_OK, font=fontStyle, width=100, height=50)
        btn_continuar.pack(side="left", padx=40)
        window_etiquetar.bind('<Return>', lambda event: btn_continuar.invoke())

        btn_cancelar = ctk.CTkButton(frame_botoes, text="Cancelar", command=lambda: window_etiquetar.destroy(
        ), fg_color=COR_BTN_CANCEL, font=fontStyle, width=100, height=50)
        btn_cancelar.pack(side="left", padx=40)
        window_etiquetar.bind('<Escape>', lambda event: btn_cancelar.invoke())


if __name__ == "__main__":  # Função base.
    app = App()
    app.mainloop()
