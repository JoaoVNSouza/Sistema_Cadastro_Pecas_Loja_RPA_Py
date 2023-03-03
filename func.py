# Importar bibliotecas necessárias.
from func_gerais import *
import numpy as np
import pyautogui as pag
import time as t
import ctypes


# Extrair dados das notas xml.
def extrair_dados(root, nsNFE, marca, tag, porcentagem):
    '''
        Objetivo: extrair os dados de cada um dos produtos contidos nos arquivos xml.

        Parameters:
            root (obj): Objeto do arquivo xml.
            nsNFE (dict): Dicionário convertido do xml.
            marca (str): Marca dos produtos.

        Returns:
            NCM, desc, cod, ref, qtd, preco, grupo_cod, grupo_nome, modify (list): Listas com todas as informações
    '''

    # Listas vazias.
    NCM, desc, cod, ref, qtd, grupo_cod, grupo_nome, preco, modify = tuple(
        [] for i in range(9))

    for item in root.findall('./ns:NFe/ns:infNFe/ns:det', nsNFE):
        ncm = int(check_none(item.find('./ns:prod/ns:NCM', nsNFE)))  # NCM
        NCM.append(ncm)

        # Quando a referência estiver junto com a descrição.
        descricao, referencia = desc_correta(
            check_none(item.find('./ns:prod/ns:xProd', nsNFE)), marca, tag)  # Descrição.
        desc.append(descricao)

        # Quando a referência NÃO estiver junto com a descrição.
        # Referência e código de algumas marcas.
        if marca == 'MARCA1' or marca == 'MARCA6' or marca == 'MARCA7':
            referencia = (check_none(item.find('./ns:prod/ns:cProd', nsNFE)))
            if marca == 'MARCA1':  # MARCA1.
                if not tag == 1:  # Seja MARCA8.
                    referencia = ref_marca1(check_none(
                        item.find('./ns:infAdProd', nsNFE)))
            elif marca == 'MARCA6':        # MARCA6.
                referencia = ref_marca6(referencia)
            else:               # MARCA7.
                referencia = referencia[1:6]
        else:
            cod.append(
                int(check_none(item.find('./ns:prod/ns:cProd', nsNFE))))

        ref.append(referencia)

        # ValorUnid.
        valorUnid = float(check_none(item.find('.ns:prod/ns:vUnTrib', nsNFE)))
        if marca == 'MARCA2':
            valorUnid *= 2
        preco_aux = calc_preco(valorUnid, porcentagem)
        preco.append(preco_aux)

        # Grupo.
        gp, grupo = verifica_grupo(descricao)
        grupo_cod.append(gp)
        grupo_nome.append(grupo)

        # Quantidade.
        quantidade = int(
            float(check_none(item.find('./ns:prod/ns:qTrib', nsNFE))))

                # Verifica se precisa modificar a descrição para caber na etiqueta com suporte até 39 carref_marca1eres.
        if (len(descricao) > 39):
            modify.append('MUDAR DESC')
        else:
            modify.append('')

        if marca == 'MARCA6':
            qtd.append(1)
            while quantidade > 1:
                NCM.append(ncm)
                desc.append(descricao)
                ref.append(referencia)
                qtd.append(1)
                preco.append(preco_aux)
                grupo_cod.append(gp)
                grupo_nome.append(grupo)
                modify.append('')
                quantidade -= 1
        else:
            qtd.append(quantidade)

    if marca == 'MARCA1' or marca == 'MARCA6' or marca == 'MARCA7':
        cod = np.nan
    elif marca != 'MARCA2':
        ref = np.nan

    return NCM, desc, cod, ref, qtd, preco, grupo_cod, grupo_nome, modify


# Extrair dados da nota Marca3.
def marca3(marca3_df):
    '''
        Objetivo: extrair os dados de cada um dos produtos contidos nos arquivos excel.

        Parameters:
            df (dataFrame): dataframe com todos os dados.

        Returns:
            NCM, desc, cod, ref, qtd, preco, grupo_cod, grupo_nome, modify (list): Listas com todas as informações
    '''

    # Extrair dados.
    cod = np.nan
    modify = []
    desc = []
    desc_aux = marca3_df['Descrição do Produto']   # Parte 1 desc.
    cor_aux = marca3_df['Descrição da Cor']        # Parte 2 desc.
    ref = marca3_df['Código Produto']
    qtd = marca3_df['Qtd. Pares']
    NCM = marca3_df['NCM']
    preco = marca3_df['Preço Sugestão'] - 0.1
    grupo_cod = 32
    grupo_nome = 'CALÇADOS MARCA3'
    tamanho = list(map(numeracao, marca3_df['Numeração']))    # Pate 3 desc.

    # Descrição correta.
    for i, _ in enumerate(desc_aux):
        cor = cor_aux[i]

        # Verifica se existe '/' na cor e pega tudo que estiver antes dela.
        if '/' in cor:
            lista = cor.split('/')
            cor = lista[0]

        # If a cor estiver mais que duas separadas por espaço.
        lista = cor.split(' ')
        if len(lista) > 1:
            cor = lista[0]

        # Remove ":" da cor.
        lista = cor.split(':')
        if len(lista) > 1:
            cor = lista[1]

        # String desc + cor.
        string = (f'{desc_aux[i]} {cor}')
        lista = tratamento_dados(string.split(' '))

        # Remover palavras com poucos carref_marca1eres sozinhos.
        lista2 = [item for item in lista if len(item) < 3 and 'I' not in item]
        [lista.remove(item) for item in lista2]

        # Unir desc + tamanho.
        lista.append(f'{tamanho[i]}')

        descricao = ' '.join(lista)
        desc.append(descricao)

        # Verifica se precisa modificar a descrição para caber na etiqueta com suporte até 39 carref_marca1eres.
        if (len(descricao) > 39):
            modify.append('MUDAR DESC')
        else:
            modify.append('')

    return NCM, desc, cod, ref, qtd, preco, grupo_cod, grupo_nome, modify


# Formatação correta da descrição.
def desc_correta(desc, marca, tag):
    '''
        Objetivo: Formatar a descrição da marca Marca1, Marca4, Marca5.

        Parameters:
            desc (str): descrição do produto sem formatação.

        Returns:
            desc (str): descrição do produto formatada.
    '''
    ref = None
    if tag == 1:  # MARCA8.
        lista = desc.split('/')  # Transforma string em lista.
        texto = lista[0].split(' ')[2]
        # lista = ['Descrição' , carref_marca1eristica, cor, tamanho]
        lista = ['BONE MASC', texto, lista[-2], lista[-1]]

    elif tag == 2:  # MARCA7.
        lista = desc.split(' ')  # Transforma string em lista.
        tamanho = lista.pop(-2)  # Retira o tamanho.
        if len(tamanho) > 2:     # Se tiver mais de 3 digítos no tamanho.
            tamanho = tamanho[1:]
        lista = tratamento_dados(lista)
        lista = desc_geral(lista)
        lista.append(tamanho)    # Adiciona o tamanho.
    else:         # Outras marcas.
        lista = desc.split(' ')  # Transforma string em lista.
        tamanho = lista.pop(-1)  # Retira o tamanho da lista.
        # Retira a referência se estiver junto desc "MARCA2".
        ref = lista[0][1:-1]

        # Fazer mudanças necessárias na descrição.
        lista = tratamento_dados(lista)  # Remover algumas PALAVRAS.
        lista = desc_geral(lista)        # Fazer algumas ações para geral.

        # Unir lista ao tamanho.
        if marca != 'MARCA6':
            lista.append(tamanho)

    # Juntar elementos da lista em uma string.
    return " ".join(lista), ref


# Função para arrumar a numeração das marca3.
def numeracao(item):
    tam = str(item).split('.')

    if tam[1] == '0':
        return tam[0]
    elif len(tam[1]) < 2:
        return f'{tam[0]}/{tam[1]}0'
    else:
        return f'{tam[0]}/{tam[1]}'


# Função para desativar o caps lock.
def caps_lock_on():
    # Verifica o estado da tecla caps lock
    return ctypes.windll.user32.GetKeyState(0x14) & 0xffff != 0
