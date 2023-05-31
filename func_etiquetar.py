# Importar bibliotecas necessárias.
from definitions import *

def abrir_opcao_etiquetar() -> None: # Abre o sistema na opção de etiquetar.
    """
    Abre o sistema na opção de etiquetar.

    Parameters: None.

    Returns: None.
    """
    
    pag.click(x=163, y=33)   # Abre o Mov. Diário.
    pag.click(x=229, y=450)  # Abre Etiquetas.
    pag.click(x=427, y=450)  # Abre Produtos.
    
    # Configurar opção por código -> Fazer antes de iniciar a automação.
    pag.press('*')
    pag.press('enter')
    pag.click(x=142, y=93)
    pag.write('Codigo')
    pag.press('enter')
    pag.press('esc')

def imprimir_etiquetas() -> None: 
    """
    Imprime as etiquetas.

    Parameters: None.

    Returns: None.
    """

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