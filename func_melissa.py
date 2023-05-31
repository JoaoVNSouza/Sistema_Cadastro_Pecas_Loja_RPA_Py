# Importar bibliotecas.
from definitions import *

def tratamento_cores_melissa(ref_df : pd.DataFrame) -> pd.DataFrame: # Tratar as cores da melissa.
    """
        Objetivo: Tratar as cores de cada item de uma referência.

        Parâmetros:
            ref_df: DataFrame com os dados da referência.

        Retorno:
            ref_df: DataFrame com os dados da referência tratados.
    """

    # Tratamento das cores separadas por ' ' ou ':'.
    for i in range(len(ref_df)): 
        ref_df.loc[i, 'Cor'] = ref_df.loc[i, 'Cor'].split(' ')[0] # Filtrar apenas 1 parte de cores compostas.
        if len(ref_df.loc[i, 'Cor'].split(':')) > 1:
            ref_df.loc[i, 'Cor'] = ref_df.loc[i, 'Cor'].split(':')[1] # Filtrar apenas 1 parte de cores compostas.

    # Tratamento das cores separadas por '/'.
    cores_possiveis_df = pd.DataFrame(columns=['Cores']) # df com todas as possibilidades de cores.
    for cor in ref_df['Cor'].unique():
        cores_possiveis_df.loc[len(cores_possiveis_df)] = [cor.split('/')[0]]
    
    # Varrer cada produto da referência.
    for i in range(len(ref_df)):
        cor = ref_df.loc[i, 'Cor'].split('/')[0]
        if cores_possiveis_df.value_counts()[cor] == 1: # Se a cor não se repeti.
            ref_df.loc[i, 'Cor'] = cor
    
    return ref_df


def tratamento_numeracao_melissa(ref_df : pd.DataFrame) -> pd.DataFrame: # Tratar as numerações da melissa.
    """
        Objetivo: Tratar a numeração de cada item de uma referência.

        Parâmetros:
            ref_df: DataFrame com os dados da referência.

        Retorno:
            ref_df: DataFrame com os dados da referência tratados.
    """
    
    for i in range(len(ref_df)): # Para cada item do dataFrame.
        tam = (str(ref_df.loc[i, 'Numeração'])).split('.')

        if tam[1] == '0': # Se for tamanho simples.
            ref_df.loc[i, 'Numeração'] = tam[0]
        elif len(tam[1]) < 2:
            ref_df.loc[i, 'Numeração'] = f'{tam[0]}/{tam[1]}0'
        else:
            ref_df.loc[i, 'Numeração'] = f'{tam[0]}/{tam[1]}'

    return ref_df


def tratamento_descricao_melissa(ref_df : pd.DataFrame) -> pd.DataFrame: # Tratar as descrições da melissa
    """
        Objetivo: Tratar a descrição de cada item de uma referência.

        Parâmetros:
            ref_df: DataFrame com os dados da referência.
        
        Retorno:
            ref_df: DataFrame com os dados da referência tratados.
    """

    for i in range(len(ref_df)): # Para cada item do DataFrame.

        # Remover palavras com qtd de caracteres menor que 3.
        lista = ref_df.loc[i, 'Descrição'].split(' ')
        lista = [i for i in lista if len(i) >= 3 or 'I' in i]
        if 'INF' in lista: # Remover palavras específicas da desc.
            lista.remove('INF')
        if 'SHINY' in lista:
            lista[lista.index('SHINY')] = 'SHI'
        
        lista.append(ref_df.loc[i, 'Cor'])          # Add cor.
        lista.append(ref_df.loc[i, 'Numeração'])    # Add tam.

        desc = ' '.join(lista) # Juntar descrição.
        ref_df.loc[i, 'Descrição'] = desc
        
        # Verifica se precisa modificar a descrição.
        if (len(desc) > 39):
            ref_df.loc[i, 'Modify'] = ('MUDAR DESCRIÇÃO')
        else:
            ref_df.loc[i, 'Modify'] = ''

    return ref_df