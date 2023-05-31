# Bibliotecas.
from definitions import *
from func_excel import *
from func_cadastro import *
from func_etiquetar import *


def gerar_excel(tela_excel: ctk) -> None:  # Função do botão gerar_excel.
    """
    Objetivo: Gerar arquivo excel com os dados extraídos do XML ou .xlsx.

    Parameters:
        window (tkinter.Tk): janela principal.

    Returns:
        None.
    """

    marca = tela_excel.combobox.get().upper()    # Marca em Maiúsculo.

    # if True:
    try:
        if marca == 'MELISSA':
            # DataFrame com os dados da melissa.
            melissa_df = pd.read_excel('nota.xlsx')
            colunas = ['NCM', 'Código Produto', 'Descrição do Produto',
                       'Descrição da Cor', 'Numeração', 'Qtd. Pares', 'Preço Sugestão']
            # Pegar apenas as colunas necessárias.
            melissa_df = melissa_df[colunas]
            NCM, desc, cod, ref, qtd, preco, grupo_cod, grupo_nome, modify = extrai_melissa(
                melissa_df)

        else:
            NCM, desc, cod, ref, qtd, preco, grupo_cod, grupo_nome, modify = extrair_dados(
                marca)
            if marca == 'ACT BONÉ':
                marca = 'ACOSTAMENTO'

        # Criando dataFrame.
        # len(cod) == 0: or all(elemento == '' for elemento in cod)
        produtos = pd.DataFrame(data={'NCM': NCM, 'Grupo': grupo_nome, 'Grupo cód.': grupo_cod, 'Código': cod,
                                'Referência': ref, 'Descrição': desc, 'Marca': marca, 'Preço': preco, 'Quantidade': qtd, 'Modificar': modify})
        # Resetar os índices a partir de 1.
        produtos.index += 1
        produtos.to_excel('Produtos.xlsx')  # Exportar dataFrame.

        # Abrir excel.tela_excel.destroy()                # Fechar janela.
        abrir_excel()

        # Verificar se alguns arquivos estão na pasta.
        if 'Falta_Cadastrar.xlsx' in os.listdir():
            os.remove('Falta_Cadastrar.xlsx')
        if 'Produtos_Codigo.xlsx' in os.listdir():
            os.remove('Produtos_Codigo.xlsx')

        pag.alert('Gerou arquivo Produtos!', 'SUCESSO!')  # Mostrar aviso.
        tela_excel.destroy()                # Fechar janela.
    except:
        pag.alert('Arquivo não encontrado ou alguma falha aparente!', 'AVISO!')


def cadastrar(tela_cadastrar: ctk) -> None:  # Função do botão cadastrar.
    """
    Objetivo: Cadastrar os produtos no sistema RETAGUARDA.

    Parameters: tela_cadastrar (ctk): janela principal.

    Returns: None.
    """

    try:
    #if True:
        produtos_df, produtos_codigo_df = ler_excel()  # Ler dados.

        VerificaTelaInicial()    # Verifica o sistema RETAGUARDA.

        pag.click(x=30, y=30)    # Abre as opções de cadastro.
        pag.click(x=50, y=135)   # Abre o menu cadastrar produtos.
        t.sleep(2)               # Pausa de 2 segundos para começar.

        # Para cada item do DataFrame.
        for i in range(1, len(produtos_df) + 1):

            # Pegar dados da tabela.
            NCM = produtos_df['NCM'][i]
            referencia = produtos_df['Referência'][i]
            descricao = produtos_df['Descrição'][i].upper()
            marca = produtos_df['Marca'][i]
            preco = produtos_df['Preço'][i]
            qtd = produtos_df['Quantidade'][i]
            grupo = produtos_df['Grupo cód.'][i]

            # Automação.
            pag.press('tab')
            pag.write(NCM)
            pag.press('tab')
            pag.write(grupo)
            apertar_tab(5)          # tab 5x.
            pag.write('n')
            pag.press('tab')
            pag.write(referencia)
            pag.press('tab')
            pag.write(descricao)
            pag.press('tab')
            pag.write(marca)
            pag.click(x=364, y=545)  # Lacuna do preço.
            pag.write(preco)
            apertar_tab(3)           # tab 3x.

            # Melissa cadastra pares e não unidade.
            if marca == 'MELISSA':
                pag.click(x=869, y=594)
                pag.write('Par')
                pag.press('enter')

            # Pega o código do produto antes de cadastrar até o final.
            pag.doubleClick(x=368, y=227)
            pag.hotkey('ctrl', 'c')
            codigo = clip.paste()[1:]

            pag.press('f4')

            if erro_NCM(tela_cadastrar):
                return  # Tratar.

            produtos_df.drop(i, axis=0, inplace=True)
            produtos_codigo_df.loc[len(
                produtos_codigo_df) + 1] = [referencia, descricao, qtd, codigo]
            produtos_df.to_excel('Falta_Cadastrar.xlsx')
            produtos_codigo_df.to_excel('Produtos_Codigo.xlsx')

        # Depois da última peça, Fecha a opção de cadastro e remove o arquivo Falta_Cadastrar.
        pag.press('esc')
        os.remove('Falta_Cadastrar.xlsx')
        pag.alert(
            'Cadastro Finalizado! Já pode utilizar o computador!', title='SUCESSO')
        tela_cadastrar.destroy()
    except:
        pag.alert('Automação cancelada, veja o arquivo \"Falta_Cadastrar\"\nCom os produtos que ainda não foram cadastrados!', title='AVISO')


def etiquetar(tela_etiquetar: ctk) -> None:  # Função etiquetar peças.
    """
    Objetivo: Etiquetar as peças no sistema RETAGUARDA.

    Parameters: tela_etiquetar (ctk): janela principal.

    Returns: None.
    """

    # if True:
    try:
        VerificaTelaInicial()    # Verifica o sistema RETAGUARDA.
        abrir_opcao_etiquetar()  # Abre a opção de etiquetar peças.
        # Ler os códigos dos produtos cadastrados.
        produtos_codigo_df = pd.read_excel(
            'Produtos_Codigo.xlsx', index_col=0, dtype=str)

        for i, codigo in enumerate(produtos_codigo_df['Código no Sistema']):
            pag.press('*')
            pag.press('enter')
            pag.write(codigo)
            pag.press('enter')
            apertar_tab(4)      # tab 4x.
            pag.write(str(produtos_codigo_df['Quantidade'][i+1]))
            apertar_tab(2)      # tab 2x.

        imprimir_etiquetas()

        pag.alert(
            'Etiquetas retirados com sucesso! Já pode utilizar o computador!', title='SUCESSO')
        tela_etiquetar.destroy()
    except:
        pag.alert('\tAutomação cancelada.\nNem todos as etiquetas foram impressas!', title='AVISO')
