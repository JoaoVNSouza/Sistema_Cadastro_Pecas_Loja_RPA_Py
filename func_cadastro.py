# Importar bibliotecas necessárias.
from definitions import *


# Função para apertar o botão tab no teclado.
def apertar_tab(num: int) -> None:
    """
    Objetivo: Apertar o botão tab no teclado.

    Parameters: num (int): número de vezes que o botão tab será apertado.

    Returns: None.
    """

    for _ in range(num):
        # Loop para apertar o botão tab o número de vezes que o usuário escolher.
        pag.press('tab')


def disable_caps_lock() -> None:  # Função para desativar o caps lock.
    """
    Desativa o caps lock.

    Paramêtros: None

    Retorno: None
    """

    # Verifica o estado da tecla caps lock.
    if ctypes.windll.user32.GetKeyState(0x14) & 0xffff != 0:
        pag.press('capslock')


# Verifica se o Retaguarda está na tela inicial.
def VerificaTelaInicial() -> None:
    """
    Verifica se o sistema RETAGUARDA está na tela inicial, caso contrário tenta voltar para ela.

    Parameters:
        None

    Returns:
        None
    """

    disable_caps_lock()  # Desativar caps_lock.

    t.sleep(1)

    # Abrir sistema.
    pag.click(x=161, y=739)
    pag.click(x=161, y=640)

    # Lista todas as imagens da pasta.
    imagens = os.listdir(f'{path_imagens}')
    imagens_retaguarda = list(
        filter(lambda img: img.startswith('retaguarda'), imagens))

    # Se estiver na tela de vendas.
    Imagefind = pag.locateOnScreen(
        f'{path_imagens}/venda.png', grayscale=True, confidence=0.9)
    if Imagefind:
        pag.click(pag.center(Imagefind))
    else:
        for item in imagens_retaguarda:
            if pag.locateOnScreen(f'{path_imagens}/{item}', grayscale=True, confidence=0.9):
                pag.press('esc')
                break


# Retorna os dataFrame necessário para cadastrar e etiquetar.
def ler_excel() -> None:
    """
    Lê os arquivos excel e retorna os dataFrame necessários para cadastrar e etiquetar.

    Parameter: None

    Returns: None
    """

    # Se existir o 'Falta_Cadastrar.xlsx'.
    if 'Falta_Cadastrar.xlsx' in os.listdir():
        produtos_df = pd.read_excel(
            'Falta_Cadastrar.xlsx', index_col=0, dtype=str)
        # Retira linhas vazias.
        produtos_df.dropna(axis=0, inplace=True, how='all')
        # Reseta os índices.
        produtos_df.reset_index(drop=True, inplace=True)
        # Adiciona 1 nos índices.
        produtos_df.index += 1
        produtos_df.to_excel('Falta_Cadastrar.xlsx')
        produtos_codigo_df = pd.read_excel(
            'Produtos_Codigo.xlsx', index_col=0, dtype=str)

    elif 'Produtos.xlsx' in os.listdir():  # Se existir o 'Produtos.xlsx'.
        produtos_df = pd.read_excel('Produtos.xlsx', index_col=0, dtype=str)
        # Retira linhas vazias.
        produtos_df.dropna(axis=0, inplace=True, how='all')
        # Reseta os índices.
        produtos_df.reset_index(drop=True, inplace=True)
        # Adiciona 1 nos índices.
        produtos_df.index += 1
        produtos_df.to_excel('Produtos.xlsx')
        produtos_df.to_excel('Falta_Cadastrar.xlsx')
        colunas = ['Referência', 'Descrição',
                   'Quantidade', 'Código no Sistema']
        produtos_codigo_df = pd.DataFrame(columns=colunas)
        produtos_codigo_df.to_excel('Produtos_Codigo.xlsx')

    return produtos_df, produtos_codigo_df


def erro_NCM(tela_cadastrar): # Tratar o erro de NCM_INVÁLIDO.
    """
    Mostrar aviso para o usuário caso o NCM não seja válido.

    Parameters: tela_cadastrar (ctk) : tela de cadastro.

    Returns: True or False (bool) : Se encontrar ou não o problema.
    """

    # Tratar o erro do NCM.
    NCM_invalido = f'{path_imagens}\\NCM_invalido.png'

    if pag.locateOnScreen(NCM_invalido):
        pag.alert(
            'NCM Inválido, Arrume o NCM na lista de \'Falta_Cadastrar.xlsx\'\n\n\t\tE tente novamente', 'CUIDADO!!!')
        tela_cadastrar.destroy()
        return True
    return False
