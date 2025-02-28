import re
import csv
from PyPDF2 import PdfReader

def extrair_texto_pdf(caminho_pdf):
    texto = ""
    with open(caminho_pdf, "rb") as f:
        leitor = PdfReader(f)
        for pagina in leitor.pages:
            texto += pagina.extract_text() + "\n"
    return texto

def extrair_dados(texto):
    produtos = []
    
    if "DESCRIÇÃO" in texto:
        texto = texto.split("DESCRIÇÃO", 1)[1]
    else:
        print("Cabeçalho 'DESCRIÇÃO' não encontrado.")
        return []
    
    linhas = [linha.strip() for linha in texto.splitlines() if linha.strip()]
    texto = "\n".join(linhas)
    
    padrao = re.compile(
        r"(?P<referencia>\[[^\]]+\])\s*"
        r"(?P<descricao>(?:(?!\[\w+\]).)*?)(?=\s*\d+[.,]\d+\s*(?:[kK][gG]|[lL]|[lL]itros?))\s*"
        r"(?P<quantidade>\d+[.,]\d+)\s*"
        r"(?P<unidade>(?:[kK][gG]|[lL]|[lL]itros?))\s+"
        r"(?P<preco>\d+[.,]\d+)\s+"
        r"(?P<impostos>IVA\s*\d+%?)\s+"
        r"(?P<amount>\d+[.,]\d+\s*€)",
        flags=re.DOTALL | re.IGNORECASE
    )
    
    matches = list(padrao.finditer(texto))
    
    if not matches:
        print("Nenhuma entrada de produto foi encontrada com o padrão definido.")
    
    for match in matches:
        referencia = match.group("referencia").strip('[]')
        descricao = " ".join(match.group("descricao").split())
        
        quantidade = match.group("quantidade").strip()
        unidade = match.group("unidade").upper()
        if unidade.startswith('L'):
            unidade = 'L'
        elif unidade.upper().startswith('K'):
            unidade = 'KG'
            
        preco = " ".join(match.group("preco").split())
        
        # Extrai apenas a percentagem do IVA
        impostos = match.group("impostos")
        percentagem = re.search(r'\d+%', impostos)
        if percentagem:
            impostos = percentagem.group()
            
        amount = " ".join(match.group("amount").split())
        
        produtos.append({
            "REFERÊNCIA": referencia,
            "DESCRIÇÃO": descricao,
            "QUANTIDADE": quantidade,
            "UNIDADE": unidade,
            "PREÇO UNITÁRIO": preco,
            "IMPOSTOS": impostos,
            "AMOUNT": amount
        })
    
    return produtos

def escrever_csv(produtos, caminho_csv):
    cabecalhos = ["REFERÊNCIA", "DESCRIÇÃO", "QUANTIDADE", "UNIDADE", "PREÇO UNITÁRIO", "IMPOSTOS", "AMOUNT"]
    with open(caminho_csv, "w", newline="", encoding="utf-8") as csvfile:
        escritor = csv.DictWriter(csvfile, fieldnames=cabecalhos)
        escritor.writeheader()
        for produto in produtos:
            escritor.writerow(produto)

def main():
    caminho_pdf = r"C:\Users\jorge\OneDrive\Ambiente de Trabalho\Cotação - N.00015.pdf"
    caminho_csv = r"C:\Users\jorge\OneDrive\Ambiente de Trabalho\cotacao.csv"
    
    texto = extrair_texto_pdf(caminho_pdf)
    produtos = extrair_dados(texto)
    
    if produtos:
        escrever_csv(produtos, caminho_csv)
        print(f"Dados exportados com sucesso para {caminho_csv}")
    else:
        print("Nenhum produto foi encontrado na cotação.")

if __name__ == "__main__":
    main()