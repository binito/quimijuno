import re

def parse_linha(linha):
    """
    Função para parsear uma linha do texto extraído do PDF e extrair os dados.
    Retorna um dicionário com os dados ou None se não houver correspondência.
    """
    # Padrão de regex ajustado para capturar os campos esperados
    padrao = re.compile(
        r"\[(?P<referencia>[^\]]+)\]\s+"            # Referência entre colchetes
        r"(?P<descricao>.*?)\s+"                    # Descrição até a quantidade
        r"(?P<quantidade>\d+[.,]\d+)\s*(?P<unidade>[kK][gG]|[lL](?:itros?)?|UN)?\s+"  # Quantidade e unidade (opcional)
        r"(?P<preco>\d+[.,]\d+)\s+"                 # Preço
        r"(?P<impostos>IVA\s*\d+%?)\s+"             # Impostos (ex.: IVA 23%)
        r"(?P<amount>\d+[.,]\d+\s*€)",             # Valor total com €
        re.IGNORECASE                               # Ignora maiúsculas/minúsculas
    )
    
    # Tenta encontrar uma correspondência na linha
    match = padrao.search(linha)
    if not match:
        return None

    # Extrai os dados do match
    dados = match.groupdict()
    referencia = dados["referencia"]
    descricao = dados["descricao"].strip()
    quantidade = float(dados["quantidade"].replace(',', '.'))  # Converte para float
    unidade = dados.get("unidade", "UN").upper()               # Unidade padrão é "UN"
    if unidade.startswith('L'):
        unidade = 'L'  # Padroniza unidade de litro
    elif unidade.startswith('K'):
        unidade = 'KG'  # Padroniza unidade de quilograma
    preco = float(dados["preco"].replace(',', '.'))            # Converte preço para float
    impostos_text = dados["impostos"]
    num_impostos = re.search(r'\d+', impostos_text)            # Extrai o número do imposto
    impostos = float(num_impostos.group()) / 100.0 if num_impostos else 0.0  # Converte para decimal
    amount = float(dados["amount"].replace('€', '').replace(',', '.').strip())  # Valor total como float

    # Retorna os dados em um dicionário
    return {
        "referencia": referencia,
        "descricao": descricao,
        "quantidade": quantidade,
        "unidade": unidade,
        "preco": preco,
        "impostos": impostos,
        "amount": amount
    }

def extrair_dados(texto):
    """
    Função para processar o texto completo e extrair os dados de todas as linhas.
    Retorna uma lista de dicionários com os produtos encontrados.
    """
    produtos = []
    
    # Divide o texto após o cabeçalho "DESCRIÇÃO", se existir
    if "DESCRIÇÃO" in texto:
        _, texto_produtos = texto.split("DESCRIÇÃO", 1)
    else:
        texto_produtos = texto

    # Divide o texto em linhas e remove linhas vazias
    linhas = [linha.strip() for linha in texto_produtos.splitlines() if linha.strip()]
    
    # Processa cada linha
    for linha in linhas:
        produto = parse_linha(linha)
        if produto:
            produtos.append(produto)
    
    # Verifica se algum produto foi encontrado
    if not produtos:
        print("Nenhum produto encontrado no texto.")
    
    return produtos

# Exemplo de uso
if __name__ == "__main__":
    # Texto de exemplo extraído de um PDF
    texto_exemplo = """
    DESCRIÇÃO
    [ABC123] Produto Exemplo 10.5 KG 20.00 IVA 23% 210.00 €
    [DEF456] Outro Produto 2.0 L 15.50 IVA 13% 31.00 €
    """
    
    # Extrai os dados do texto
    produtos = extrair_dados(texto_exemplo)
    
    # Imprime os resultados
    for produto in produtos:
        print(produto)