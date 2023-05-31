# Importar bibliotecas necessárias.
from definitions import *
from func_melissa import *
from func_xml import *


def extrair_dados(marca: str) -> Tuple[List[int], List[str], List[int], List[str], List[int], List[float], List[int], List[str], List[str]]:
    """
    Objetivo: Extrair dados das notas xml.

    Parâmetros:
        marca: Marca do produto.

    Retorno:
        Uma tupla contendo os seguintes elementos:
        - NCM: Lista com os NCMs.
        - desc: Lista com as descrições.
        - cod: Lista com os códigos.
        - ref: Lista com as referências.
        - qtd: Lista com as quantidades.
        - preco: Lista com os preços.
        - grupo_cod: Lista com os códigos dos grupos.
        - grupo_nome: Lista com os nomes dos grupos.
        - modify: Lista com as modificações.
    """

    # Criando objeto com o XML.
    tree = ET.parse('nota.xml')
    root = tree.getroot()
    nsNFE = {'ns': "http://www.portalfiscal.inf.br/nfe"}  # Dicionário.

    # Listas vazias.
    NCM, desc, cod, ref, qtd, grupo_cod, grupo_nome, preco, modify = tuple(
        [] for i in range(9))

    # Para cada item da nota.
    for item in root.findall('./ns:NFe/ns:infNFe/ns:det', nsNFE):
        # Adiciona o NCM na lista.
        ncm = int(check_none(item.find('./ns:prod/ns:NCM', nsNFE)))
        NCM.append(ncm)

        # Retorna a descrição correta. Se for 'ILICITO' retorna a referencia também.
        # Descrição.
        descricao, referencia, codigo = descricao_correta(
            check_none(item.find('./ns:prod/ns:xProd', nsNFE)), marca)
        desc.append(descricao)

        # Retorna a referência correta para as demais marcas.
        if marca != 'ILICITO':
            codigo, referencia = referencia_correta(marca, item, nsNFE)

        ref.append(referencia)
        cod.append(codigo)

        # ValorUnid.
        valorUnid = float(check_none(item.find('.ns:prod/ns:vUnTrib', nsNFE)))
        if marca == 'ILICITO':
            valorUnid *= 2
        preco_aux = calc_preco(valorUnid)  # Preço de venda.
        preco.append(preco_aux)

        # Grupo.
        # Retorna o Nome e código do grupo.
        gp, grupo = verifica_grupo(descricao)
        grupo_cod.append(gp)
        grupo_nome.append(grupo)

        # Quantidade.
        quantidade = int(
            float(check_none(item.find('./ns:prod/ns:qTrib', nsNFE))))
        if marca == 'SLY':
            qtd.append(1)
            while quantidade > 1:
                NCM.append(ncm)
                desc.append(descricao)
                cod.append(codigo)
                ref.append(referencia)
                qtd.append(1)
                preco.append(preco_aux)
                grupo_cod.append(gp)
                grupo_nome.append(grupo)
                modify.append('')
                quantidade -= 1
        else:
            qtd.append(quantidade)

        # Verifica se precisa modificar a descrição.
        if (len(descricao) > 39):
            modify.append('MUDAR DESCRIÇÃO')
        else:
            modify.append('')

    return NCM, desc, cod, ref, qtd, preco, grupo_cod, grupo_nome, modify


def extrai_melissa(melissa_df_copy: pd.DataFrame) -> tuple[List[str], List[str], str, List[str], List[float], List[float], int, str, List[object]]:
    """
    Objetivo: Extrair dados da nota melissa.

    Parâmetros:
        melissa_df: DataFrame com os dados da nota melissa.

    Retorno:
        Uma tupla contendo os seguintes elementos:
        - NCM: Lista com os NCMs.
        - desc: Lista com as descrições.
        - cod: Uma string vazia.
        - ref: Lista com as referências.
        - qtd: Lista com as quantidades.
        - preco: Lista com os preços.
        - grupo_cod: Um inteiro representando o código do grupo.
        - grupo_nome: Uma string com o nome do grupo.
        - modify: Lista com as modificações.
    """

    # Filtrar e renomear as colunas necessárias.
    melissa_df = melissa_df_copy.copy()
    melissa_df.rename(columns={'Código Produto': 'Referência', 'Descrição do Produto': 'Descrição',
                               'Descrição da Cor': 'Cor', 'Qtd. Pares': 'Qtd', 'Preço Sugestão': 'Preço'}, inplace=True)

    completo_df = pd.DataFrame(columns=melissa_df.columns)
    for ref in melissa_df['Referência'].unique():  # Para cada referência.

        # Filtrar produtos de uma ref.
        ref_df = melissa_df.loc[melissa_df['Referência'] == ref]
        ref_df.reset_index(inplace=True)      # Resetar índices.
        ref_df = ref_df.drop('index', axis=1)  # Remover coluna index.

        # Tratamento dos dados.
        ref_df = tratamento_cores_melissa(ref_df)
        ref_df = tratamento_numeracao_melissa(ref_df)
        ref_df = tratamento_descricao_melissa(ref_df)

        # Novo df completo com todos os produtos tratados.
        completo_df = pd.concat([completo_df, ref_df], ignore_index=True)

    completo_df['Preço'] = completo_df['Preço'] - 0.1  # Mudar o preço.
    grupo_cod = 32
    grupo_nome = 'CALÇADOS MELISSA'
    cod = ''

    return completo_df['NCM'], completo_df['Descrição'], cod, completo_df['Referência'], completo_df['Qtd'], completo_df['Preço'], grupo_cod, grupo_nome, completo_df['Modify']


def abrir_excel() -> None:  # Abre o excel com os produtos gerados da nota.
    """
    Objetivo: Abrir o excel com os produtos gerados da nota.

    Parâmetros:
        Nenhum.

    Retorno:
        Nenhum.
    """

    t.sleep(0.5)
    pag.hotkey('win', 'r')
    pag.write(f'{path_arquivos}\Produtos.xlsx')
    pag.press('enter')
    t.sleep(4)

    # Expandir colunas.
    pag.click(x=600, y=300, clicks=2)  # Clicar no meio da tela.
    pag.hotkey('ctrl', 't')  # Selecionar todas as linhas.
    t.sleep(0.5)
    pag.click(x=958, y=145)  # Clica em 'Formatar'.
    t.sleep(0.5)
    pag.click(x=987, y=293)  # Clica em 'AutoAjuste'.

    # Definir coluna 'F' como texto.
    pag.click(x=368, y=253)  # Clica na coluna 'F'.
    pag.click(x=622, y=102)  # Clica na opção de mudar o tipo.
    pag.write('Texto')
    pag.press('enter')
