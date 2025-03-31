# Importar as bibliotecas necessárias
import xml.etree.ElementTree as ET  # Para processar o XML
import pandas as pd                 # Para criar DataFrames e exportar para Excel
import customtkinter as ctk
from tkinter import filedialog, messagebox

# Função para extrair dados de um elemento, retornando um dicionário achatado
def extract_element_data(element, ns, prefix=""):
    data = {}
    if element is None:
        return data
    # Extrai os atributos ou texto direto se não tiver sub-elementos
    if list(element):  # Se o elemento tiver sub-elementos
        for child in element:
            tag = child.tag.split('}')[1]
            # Se o filho tiver também sub-elementos, chama recursivamente
            if list(child):
                sub_data = extract_element_data(child, ns, prefix=f"{prefix}{tag}_")
                data.update(sub_data)
            else:
                data[f"{prefix}{tag}"] = child.text
    else:
        data[prefix.rstrip("_")] = element.text
    return data

# Classe da GUI
class SaftAnalyzerGUI:
    def __init__(self):
        self.app = ctk.CTk()
        self.app.title("SAFT Analyzer")
        self.app.geometry("600x400")
        
        # Configuração do grid
        self.app.grid_columnconfigure(1, weight=1)
        
        # Seleção do ficheiro SAFT
        ctk.CTkLabel(self.app, text="Ficheiro SAFT:").grid(row=0, column=0, padx=10, pady=10)
        self.saft_entry = ctk.CTkEntry(self.app, width=400)
        self.saft_entry.grid(row=0, column=1, padx=10, pady=10)
        ctk.CTkButton(self.app, text="Procurar", command=self.browse_saft).grid(row=0, column=2, padx=10, pady=10)
        
        # Seleção do ficheiro Excel destino
        ctk.CTkLabel(self.app, text="Excel destino:").grid(row=1, column=0, padx=10, pady=10)
        self.excel_entry = ctk.CTkEntry(self.app, width=400)
        self.excel_entry.grid(row=1, column=1, padx=10, pady=10)
        ctk.CTkButton(self.app, text="Procurar", command=self.browse_excel).grid(row=1, column=2, padx=10, pady=10)
        
        # Botão para processar o SAFT
        ctk.CTkButton(self.app, text="Processar SAFT", command=self.process_saft).grid(row=2, column=0, columnspan=3, pady=20)
        
    def browse_saft(self):
        filename = filedialog.askopenfilename(filetypes=[("XML files", "*.xml")])
        if filename:
            self.saft_entry.delete(0, 'end')
            self.saft_entry.insert(0, filename)
            
    def browse_excel(self):
        filename = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                              filetypes=[("Excel files", "*.xlsx")])
        if filename:
            self.excel_entry.delete(0, 'end')
            self.excel_entry.insert(0, filename)
    
    def process_saft(self):
        saft_path = self.saft_entry.get()
        excel_path = self.excel_entry.get()
        
        if not saft_path or not excel_path:
            messagebox.showerror("Erro", "Por favor selecione o ficheiro SAFT e o destino Excel")
            return
            
        try:
            # Definição do namespace e processamento do XML
            ns = {"saft": "urn:OECD:StandardAuditFile-Tax:PT_1.04_01"}
            tree = ET.parse(saft_path)
            root = tree.getroot()
            
            # === EXTRAIR DADOS DO CABEÇALHO ===
            header = root.find("saft:Header", ns)
            header_data = extract_element_data(header, ns)
            df_header = pd.DataFrame([header_data])
            
            # === EXTRAIR DADOS DOS CLIENTES (MasterFiles -> Customer) ===
            master_files = root.find("saft:MasterFiles", ns)
            customers = master_files.findall("saft:Customer", ns)
            customer_list = []
            for customer in customers:
                customer_data = extract_element_data(customer, ns)
                customer_list.append(customer_data)
            df_customers = pd.DataFrame(customer_list)
            
            # === EXTRAIR DADOS DA TABELA DE IVA (MasterFiles -> TaxTable) ===
            tax_table = master_files.find("saft:TaxTable", ns)
            tax_entries = tax_table.findall("saft:TaxTableEntry", ns)
            tax_list = []
            for entry in tax_entries:
                entry_data = extract_element_data(entry, ns)
                tax_list.append(entry_data)
            df_tax = pd.DataFrame(tax_list)
            
            # === EXTRAIR DADOS DAS FATURAS DE VENDAS (SourceDocuments -> SalesInvoices) ===
            source_docs = root.find("saft:SourceDocuments", ns)
            sales_invoices = source_docs.find("saft:SalesInvoices", ns)
            invoice_elements = sales_invoices.findall("saft:Invoice", ns)
            invoice_list = []
            for inv in invoice_elements:
                inv_data = {}
                # Extrair campos básicos e elementos aninhados usando a função auxiliar
                inv_data.update(extract_element_data(inv, ns))
                # Se pretender extrair informações de linhas (Line) individualmente,
                # pode também iterar sobre esses elementos e guardar os resultados numa lista
                lines = inv.findall("saft:Line", ns)
                line_list = []
                for line in lines:
                    line_list.append(extract_element_data(line, ns, prefix="Line_"))
                # Armazenar os dados das linhas como uma string ou manter como lista (dependendo da análise)
                inv_data["Lines"] = line_list
                invoice_list.append(inv_data)
            df_invoices = pd.DataFrame(invoice_list)
            
            # === EXPORTAR OS DADOS PARA UM FICHEIRO EXCEL ===
            with pd.ExcelWriter(excel_path) as writer:
                df_header.to_excel(writer, sheet_name="Header", index=False)
                df_customers.to_excel(writer, sheet_name="Clientes", index=False)
                df_tax.to_excel(writer, sheet_name="TaxTable", index=False)
                # Para as faturas, pode ser interessante guardar os dados brutos ou processados
                df_invoices.to_excel(writer, sheet_name="SalesInvoices", index=False)
            
            messagebox.showinfo("Sucesso", f"Ficheiro Excel criado com sucesso em:\n{excel_path}")
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao processar o SAFT:\n{str(e)}")
    
    def run(self):
        self.app.mainloop()

if __name__ == "__main__":
    app = SaftAnalyzerGUI()
    app.run()
