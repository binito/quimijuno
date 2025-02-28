import tkinter as tk
from tkinter import filedialog, messagebox
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

def processar():
    pdf_path = filedialog.askopenfilename(
        title="Selecione o arquivo PDF da cotação",
        filetypes=[("Arquivos PDF", "*.pdf")]
    )
    if not pdf_path:
        return

    csv_path = filedialog.asksaveasfilename(
        title="Salvar CSV como",
        defaultextension=".csv",
        filetypes=[("Arquivos CSV", "*.csv")]
    )
    if not csv_path:
        return

    try:
        texto = extrair_texto_pdf(pdf_path)
        produtos = extrair_dados(texto)
        if produtos:
            escrever_csv(produtos, csv_path)
            messagebox.showinfo("Sucesso", f"Dados exportados com sucesso para:\n{csv_path}")
        else:
            messagebox.showwarning("Atenção", "Nenhum produto foi encontrado na cotação.")
    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro: {e}")

def main():
    root = tk.Tk()
    root.title("Processamento de Cotação PDF")
    root.geometry("400x200")
    
    label = tk.Label(root, text="Clique no botão para processar a cotação:")
    label.pack(pady=20)
    
    process_button = tk.Button(root, text="Processar Cotação", command=processar)
    process_button.pack(pady=10)
    
    root.mainloop()

if __name__ == "__main__":
    main()