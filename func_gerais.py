# Importar bibliotecas necessárias.
import pandas as pd

# Constantes.
absolute_path = r'C:\user\directory\Arquivos'


# Verificação caso os dados estejam vazios.
def check_none(var):
    '''
        Objetivo: Verificar se o conteúdo de uma tag xml contém informação ou é vazio.

        Parameters:
            var (str or None): conteúdo de uma taga xml.

        Returns:
            '' or text: retorna vazio ou o conteúdo presente na tag.
    '''
    if var == None:
        return ''
    else:
        return var.text


# Função para tratamento de dados.
def tratamento_dados(lista):
    '''
        Objetivo: Realizar o tratamento da descrição dos produtos, removendo e alterando palavras que não necessárias para o cadastro.

        Parameters:
        lista (list): descrição do produto separada por espaços.

        Return:
        lista (list): lista após o tratamento de dados.
    '''


    # Remover palavras da descrição.
    df = pd.read_excel('Remover da descrição.xlsx')
    for palavra in df['Palavras']:
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
    if 'MARCA1' in lista:
        lista[lista.index('MARCA1')] = 'ref_marca1'
    if 'BOOT' in lista:
        lista[lista.index('BOOT')] = 'CALCA'
    if 'JEGGING' in lista:
        lista[lista.index('JEGGING')] = 'CALCA'

    # Remover palavras que contenham uma determinada sigla.
    [lista.pop(i) for i, item in enumerate(lista) if 'COL' in item]

    # Remover palavras que contenham uma determinada sigla.
    [lista.pop(i) for i, item in enumerate(lista) if 'TRADI' in item]

    return lista


# Remover itens na lista com poucos caracteres.
def desc_geral(lista):

    # Alterações globais.
    if ('T-SHIRT' in lista or 'SHIRT' in lista or 'TSHIRT' in lista):
        if ('MASC' not in lista):
            try:
                lista[lista.index('T-SHIRT')] = 'BLUSA'
            except:
                try:
                    lista[lista.index('SHIRT')] = 'BLUSA'
                except:
                    lista[lista.index('TSHIRT')] = 'BLUSA'
        else:
            try:
                lista[lista.index('T-SHIRT')] = 'CASACO'
            except:
                try:
                    lista[lista.index('SHIRT')] = 'CASACO'
                except:
                    lista[lista.index('TSHIRT')] = 'CASACO'

    # Se existir cores com o 'C/'.
    if 'C/' in lista:
        lista = lista[: lista.index('C/')]
    # Se existir cores com o 'C/'.
    if 'S/' in lista:
        lista = lista[: lista.index('S/')]

    # Separar cores dos números no final da descrição.
    # Só vai ser feito caso seja 'MARCA2', pois outros não tem essa '('.
    aux = lista[-1]
    lista_aux = aux.split('(')
    if len(lista_aux) > 1:
        lista[-1] = lista_aux[0]
    else:
        pass

    # Remover palavras com poucos caracteres sozinhos, números ou símbolos da descrição.
    [lista.pop(i) for i, item in enumerate(lista) if (((len(item) < 3) and (
        item != 'MC' and item != 'ML' and item != 'VD')) or not item.isalpha())]

    # Verificar se falta o 'FEM' depois dos tipo de produto feminino.
    if ('FEM' not in lista) or ('MASC' not in lista):
        if ('CALCA' in lista or 'VESTIDO' in lista or 'BLUSA' in lista or 'CROPPED' in lista or 'SAIA' in lista or 'TOP' in lista or 'BLAZER' in lista or 'CAMISA' in lista or 'SHORTS' in lista):
            try:
                if lista[1] != 'FEM':
                    lista.insert(1, 'FEM')
            except:
                lista.append('FEM')
    if ('BLUSA' in lista or 'CROPPED' in lista) and ('MC' not in lista and 'ML' not in lista):
        lista.insert(2, 'MC')

    return lista


# Faz o cálculo do preço de venda dos produtos.
def calc_preco(valorUnid, porcentagem):
    valor = round((valorUnid * porcentagem) + valorUnid)

    lista = list(str(valor))
    while (valor % 10) != 0:
        if (lista[-1] == '1') or (lista[-1] == '2'):
            valor -= 1
        else:
            valor += 1

    valor -= 0.1

    return valor


# Ajusta a referência da MARCA1.
def ref_marca1(var):
    '''
        Objetivo: formatar a referência para o formato utilizado pela loja.

        Parameters:
            var (str): Referência da peça.

        Returns:
            ref (str): Referência da peça formatada.
    '''
    lista = var.split(' ')  # Dividir a referência em uma lista.

    # Filtra apenas os números da referência.
    if len(lista) == 1:  # Somente a referência.
        ref = lista[0]
    # Se a 1º letra do 1º for número, essa é a referência.
    elif lista[0][0].isnumeric():
        ref = lista[0]
    else:
        ref = lista[-1]  # O último elemento da lista é a referência.

     # Remover 'M' no final mas deixar outras letras.
    if ref[-1] == 'M':
        ref = ref[: -1]

    return ref


# Ajusta a referência da MARCA6.
def ref_marca6(var):
    return int(str(var).replace('.', ''))


# Verifica se camiseta é infantil.
def infantil(desc):
    '''
        Objetivo: Verificar se um produto contendo 'CAMISETA' na descrição é do grupo Infantil.

        Parameters:
            desc (str): Descrição do produto.

        Returns:
            True: Caso seja infantil.
    '''

    if '02' in desc or '04' in desc or '06' in desc or '08' in desc or '10' in desc or '12' in desc or '14' in desc or '16' in desc or '2' in desc or '4' in desc or '6' in desc or '8' in desc:
        return True
    else:
        return False


# Reponsável por verificar a qual grupo pertence aquele produto.
def verifica_grupo(desc):
    '''
        Objetivo: Baseado na descrição do produto formatado, verifica qual o grupo o produto pertence.

        Parameters:
            desc (str): Descrição do produto.

        Returns:
            cod (int): Código do grupo.
            nome grupo (str): nome do grupo.
    '''

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
        elif infantil(desc):    # Infantil.
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
