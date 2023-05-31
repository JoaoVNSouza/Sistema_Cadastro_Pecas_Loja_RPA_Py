from definitions import *


def check_none(var):  # Verificação caso os dados estejam vazios.
    """
    Objetivo: Verificar se o conteúdo de uma tag xml contém informação ou é vazio.

    Parameters:
        var (str or None): conteúdo de uma taga xml.

    Returns:
        '' or text: retorna vazio ou o conteúdo presente na tag.
    """
    if var == None:
        return ''
    else:
        return var.text


# Melhora a descrição, removendo e ajustando palavras.
def tratamento_dados(lista: List[str]) -> List[str]:
    """
    Função para remover palavras ou ajustar algumas na descrição.

    Parâmetros:
        lista (list): Lista de strings contendo as palavras a serem tratadas.

    Retorno:
        lista (list): Lista de strings com as palavras tratadas.
    """

    # Remover palavras da descrição.
    words_rm_df = pd.read_excel('Palavras para REMOVER da descrição.xlsx')
    for palavra in words_rm_df['Palavras']:
        if palavra in lista:
            lista.remove(palavra)

    # Mudar palavras da descrição.
    if 'MASC.' in lista:
        lista[lista.index('MASC.')] = 'MASC'
    if 'FEM.' in lista:
        lista[lista.index('FEM.')] = 'FEM'
    if 'BE.' in lista:
        lista[lista.index('BE.')] = 'BERMUDA'
    if 'CA.' in lista:
        lista[lista.index('CA.')] = 'CALCA'
    if '(ML)' in lista:
        lista[lista.index('(ML)')] = 'ML'
    if '(MC)' in lista:
        lista[lista.index('(MC)')] = 'MC'
    if 'ACOSTAMENTO' in lista:
        lista[lista.index('ACOSTAMENTO')] = 'ACT'
    if 'BOOT' in lista:
        lista[lista.index('BOOT')] = 'CALCA'
    if 'JEGGING' in lista:
        lista[lista.index('JEGGING')] = 'CALCA'

    # Mudar palavras da descrição.
    palavras = ['T-SHIRT', 'SHIRT', 'TSHIRT']
    for palavra in palavras:
        if palavra in lista:
            if 'MASC' not in lista:  # Fem
                lista[lista.index(palavra)] = 'BLUSA'
            else:                   # Masc.
                lista[lista.index(palavra)] = 'CASACO'
            break

    # Se existir cores com o 'C/' ou 'S/'.
    if 'C/' in lista:
        lista = lista[: lista.index('C/')]
    if 'S/' in lista:
        lista = lista[: lista.index('S/')]

    # Separar a cor dos números do final da descrição da 'ILICITO'. Pois as cores podem estar juntas de '('.
    cor = lista[-1]
    lista_cor = cor.split('(')
    if len(lista_cor) > 1:
        lista[-1] = lista_cor[0]

    # Remover palavras com poucos caracteres, números ou símbolos da descrição.
    lista = [item for item in lista if ((len(item) >= 3) or (
        item == 'MC' or item == 'ML' or item == 'VD')) and item.isalpha()]

    # Remover palavras com 'COLOR' ou 'AMARR' dentro da lista.
    lista = [item for item in lista if 'COLOR' not in item]
    lista = [item for item in lista if 'AMARR' not in item]

    if 'FEM' not in lista and 'MASC' not in lista:
        try:
            if lista[1] != 'FEM':
                lista.insert(1, 'FEM')
        except:  # SLY.
            lista.append('FEM')

    # Adiciona 'MC' caso não tenha na descrição.
    try:
        if ('BLUSA' in lista[0] or 'CROPPED' in lista[0]) and ('MC' not in lista[2] and 'ML' not in lista[2]):
            lista.insert(2, 'MC')
    except:  # SLY.
        lista.append('MC')

    return lista


def calc_preco(valorUnid):  # Cálculo do preço de venda dos produtos.
    """
    Objetivo: Calcular o preço de venda dos produtos.

    Parameters:
        valorUnid (float): Valor unitário do produto.

    Returns:
        valor (float): Valor de venda do produto.
    """

    valor = round((valorUnid * 1.3) + valorUnid)

    lista = list(str(valor))
    while (valor % 10) != 0:
        if (lista[-1] == '1') or (lista[-1] == '2'):
            valor -= 1
        else:
            valor += 1

    valor -= 0.1  # Ficar com final 0.9

    return valor


def ref_act(referencia):  # Ajusta a referência da acostamento.
    '''
    Objetivo: formata a referência corretamento para o perfil utilizado pela loja.

    Parameters:
        referencia (str): Referência da peça.

    Returns:
        referencia (str): Referência da peça formatada.
    '''

    lista = referencia.split(' ')  # Dividir a referência em uma lista.
    # O primeiro da lista é a referência.
    if len(lista) == 1 or lista[0][0].isnumeric():
        referencia = lista[0]
    else:                         # O último elemento da lista é a referência.
        referencia = lista[-1]

    if referencia[-1] == 'M':
        # Remover 'M' se estiver no final da referência.
        referencia = referencia[: -1]

    return referencia


def ref_sly(referencia):  # Ajusta a referência da sly.
    """
    Objetivo: formata a referência corretamento para o perfil utilizado pela loja.

    Parameters:
        referencia (str): Referência da peça.

    Returns:
        referencia (str): Referência da peça formatada.
    """

    return int(str(referencia).replace('.', ''))


def camiseta_infantil(desc):  # Verifica se a descrição camiseta é infantil.
    """
    Objetivo: Verificar se um produto contendo 'CAMISETA' na descrição é do grupo Infantil.

    Parameters:
        desc (str): Descrição do produto.

    Returns:
        bool: True se for infantil, False se não for.
    """

    if '02' in desc or '04' in desc or '06' in desc or '08' in desc or '10' in desc or '12' in desc or '14' in desc or '16' in desc or '2' in desc or '4' in desc or '6' in desc or '8' in desc:
        return True
    else:
        return False


# Retorna o grupo de um produto baseado na sua descrição.
def verifica_grupo(desc):
    """
    Objetivo: Baseado na descrição do produto formatado, verifica qual o grupo o produto pertence.

    Parameters:
        desc (str): Descrição do produto.

    Returns:
        cod (int): Código do grupo.
        nome grupo (str): nome do grupo.
    """

    if 'ACESSÓRIOS' in desc:
        return 24, 'ACESSÓRIOS'
    elif 'BERMUDA' in desc:
        if 'FEM' in desc:   # Fem.
            return 22, 'BERMUDAS FEM'
        else:               # Masc.
            return 16, 'BERMUDAS MASC'
    elif 'BIQUINI' in desc or 'MAIO' in desc:
        return 23, 'BIQUINIS'
    elif 'BLAZER' in desc or 'PALETO' in desc:
        return 12, 'BLAZER'
    elif 'MOLETOM' in desc or 'BLUSAO' in desc:
        return 29, 'MOLETOM'
    elif 'CALCA' in desc or 'CALÇA' in desc:
        if 'FEM' in desc:   # Calça fem.
            return 4, 'CALÇAS FEM'
        else:               # Calça masc.
            return 3, 'CALÇAS MASC'
    elif 'BLUSA' in desc or 'TOP' in desc or 'REGATA' in desc or 'BODY' in desc or 'CROPPED' in desc:
        return 11, 'BLUSAS FEM'
    elif 'BOLSA' in desc:
        return 25, 'BOLSAS'
    elif 'BONE' in desc:
        return 20, 'BONÉS'

    elif 'CAMISA' in desc:
        if 'MASC' in desc:  # Camisa masc.
            return 17, 'CAMISAS MASC'
        else:               # Camisa fem.
            return 11, 'BLUSAS FEM'
    elif 'POLO' in desc:
        if 'FEM' in desc:    # Polo fem.
            return 11, 'BLUSAS FEM'
        else:                # Polo masc.
            return 5, 'CAMISETAS MASC'
    elif 'CAMISETA' in desc:
        if 'G2' in desc:
            return 5, 'CAMISETAS MASC'
        elif camiseta_infantil(desc):    # Infantil.
            return 8, 'CAMISETAS MASC INFANTIL'
        else:                   # Adulto.
            return 5, 'CAMISETAS MASC'
    elif 'CAMISETE' in desc:
        return 10, 'CAMISETES FEM'
    elif 'CARDIGAN' in desc:
        return 33, 'CARDIGAM'
    elif 'CARTEIRA' in desc:
        return 21, 'CARTEIRA'
    elif 'SAIA' in desc:
        if 'PRAIA' in desc:  # Saia de Praia.
            return 30, 'SAIAS DE PRAIA'
        else:               # Saia.
            return 14, 'SAIAS'
    elif 'CASACO' in desc or 'TRICOT' in desc:
        if 'FEM' in desc:   # Casaco fem.
            return 9, 'CASACOS FEM'
        else:               # Casaco masc.
            return 6, 'CASACOS MASC'

    elif 'CHINELO' in desc:
        return 26, 'CHINELOS'
    elif 'CINTO' in desc:
        return 19, 'CINTOS'
    elif 'CONJUNTO' in desc:
        return 27, 'CONJUNTOS'
    elif 'SHORT' in desc:
        if 'FEM' in desc:   # Short fem.
            return 18, 'SHORTS FEM'
        else:               # Bermuda masc.
            return 16, 'BERMUDAS MASC'
    elif 'CUECA' in desc or 'BOXER' in desc:
        return 31, 'CUECA'
    elif 'JAQUETA' in desc:
        return 28, 'JAQUETAS'
    elif 'MACACAO' in desc:
        return 15, 'MACACAO'
    elif 'VESTIDO' in desc:
        return 7, 'VESTIDOS'
    else:
        return 11, 'BLUSAS FEM'  # CASO GERAL, retorna blusa.


# Formatação correta da descrição.
def descricao_correta(desc: str, marca: str) -> str:
    """
    Objetivo: Formatar corretamente a descrição recebida e se for Ilicito retorna a referência.

    Parâmetros:
        desc: Descrição do produto.
        marca: Marca do produto.

    Retorno:
        descricao: Descrição formatada.
        ref: Referência do produto.
    """

    ref = ''  # Valor padrão da referência.
    cod = ''  # Valor padrão da referência.

    # Boné ACT.
    if marca == 'ACT BONÉ':
        lista = desc.split('/')
        texto = lista[0].split(' ')[2]
        lista = ['BONE MASC', texto, lista[-2], lista[-1]]

    # Outros.
    else:
        lista = desc.split(' ')         # Transforma string em lista.
        if marca == 'YESKLA':
            tamanho = lista.pop(-2)     # Retira o tamanho.
            # Se tiver mais de 3 digítos no tamanho.
            if len(tamanho) > 2:
                tamanho = tamanho[1:]   # Ajusta o tamanho.
        else:
            tamanho = lista.pop(-1)     # Retira o tamanho.

        # Retira a referência junto da desc caso seja "Ilicito".
        ref = lista[0][1:-1]

        # Fazer mudanças necessárias na descrição.
        # Remover, ajustar ou adicionar algumas palavras.
        lista = tratamento_dados(lista)

        # Unir o tamanho a lista.
        if marca != 'SLY':
            lista.append(tamanho)

    descricao = " ".join(lista)  # Lista -> string.

    return descricao, ref, cod


# Formatação correta da referencia.
def referencia_correta(marca: str, item, nsNFE) -> Tuple[str, str]:
    """
    Retorna o código e a referência correta com base na marca especificada.

    Paramêtros:
        marca (str): A marca do item.
        item: O item a ser verificado.
        nsNFE: O namespace da NFE.

    Retorno:
        codigo (str): O código do item.
        referencia (str): A referência do item.
    """

    codigo = ''        # Valor padrão.
    referencia = ''    # Valor padrão.

    if marca == 'ACOSTAMENTO':
        referencia = check_none(item.find('./ns:infAdProd', nsNFE))
        referencia = ref_act(referencia)
    elif marca == 'SLY' or marca == 'ACT BONÉ' or marca == 'YESKLA':
        referencia = check_none(item.find('./ns:prod/ns:cProd', nsNFE))
        if marca == 'SLY':        # Sly.
            referencia = ref_sly(referencia)
        elif marca == 'YESKLA':   # Yeskla.
            referencia = referencia[1:6]
    elif marca == 'INDEX' or marca == 'CHOPPER':
        codigo = int(check_none(item.find('./ns:prod/ns:cProd', nsNFE)))

    return codigo, referencia
