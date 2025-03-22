import csv
import os
import tkinter as tk
from tkinter import filedialog, messagebox
import customtkinter as ctk
import pyvat
import threading
import pandas as pd
from datetime import datetime
import time  # Importa o módulo time
import requests  # Importa a biblioteca requests

class VATValidatorApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        # Configuração básica da janela
        self.title("Validador de VAT")
        self.geometry("700x500")
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(3, weight=1)
        
        # Variáveis
        self.input_file = tk.StringVar()
        self.output_file = tk.StringVar()
        self.vat_column = tk.StringVar(value="vat")
        self.country_column = tk.StringVar(value="country")
        self.processing = False
        self.results = []
        
        # Widgets
        self.create_widgets()
    
    def create_widgets(self):
        # Frame de arquivos
        file_frame = ctk.CTkFrame(self)
        file_frame.grid(row=0, column=0, padx=20, pady=(20, 10), sticky="ew")
        file_frame.grid_columnconfigure(1, weight=1)
        
        # Input file
        ctk.CTkLabel(file_frame, text="Arquivo de entrada:").grid(row=0, column=0, padx=10, pady=10, sticky="w")
        ctk.CTkEntry(file_frame, textvariable=self.input_file).grid(row=0, column=1, padx=10, pady=10, sticky="ew")
        ctk.CTkButton(file_frame, text="Procurar", command=self.browse_input).grid(row=0, column=2, padx=10, pady=10)
        
        # Output file
        ctk.CTkLabel(file_frame, text="Arquivo de saída:").grid(row=1, column=0, padx=10, pady=10, sticky="w")
        ctk.CTkEntry(file_frame, textvariable=self.output_file).grid(row=1, column=1, padx=10, pady=10, sticky="ew")
        ctk.CTkButton(file_frame, text="Procurar", command=self.browse_output).grid(row=1, column=2, padx=10, pady=10)
        
        # Frame de opções
        options_frame = ctk.CTkFrame(self)
        options_frame.grid(row=1, column=0, padx=20, pady=10, sticky="ew")
        options_frame.grid_columnconfigure(1, weight=1)
        
        # Nome das colunas
        ctk.CTkLabel(options_frame, text="Coluna VAT:").grid(row=0, column=0, padx=10, pady=10, sticky="w")
        ctk.CTkEntry(options_frame, textvariable=self.vat_column).grid(row=0, column=1, padx=10, pady=10, sticky="ew")
        
        ctk.CTkLabel(options_frame, text="Coluna País:").grid(row=1, column=0, padx=10, pady=10, sticky="w")
        ctk.CTkEntry(options_frame, textvariable=self.country_column).grid(row=1, column=1, padx=10, pady=10, sticky="ew")
        
        # Botão de execução
        self.process_button = ctk.CTkButton(self, text="Validar VATs", command=self.start_validation, fg_color="green")
        self.process_button.grid(row=2, column=0, padx=20, pady=10, sticky="ew")
        
        # Área de log
        log_frame = ctk.CTkFrame(self)
        log_frame.grid(row=3, column=0, padx=20, pady=(10, 20), sticky="nsew")
        log_frame.grid_columnconfigure(0, weight=1)
        log_frame.grid_rowconfigure(0, weight=1)
        
        self.log_text = ctk.CTkTextbox(log_frame, height=200)
        self.log_text.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")
        
        # Barra de progresso
        self.progress_bar = ctk.CTkProgressBar(self)
        self.progress_bar.grid(row=4, column=0, padx=20, pady=(0, 20), sticky="ew")
        self.progress_bar.set(0)
    
    def browse_input(self):
        filename = filedialog.askopenfilename(
            title="Selecione o arquivo CSV de entrada",
            filetypes=[("Arquivos CSV", "*.csv"), ("Todos os arquivos", "*.*")]
        )
        if filename:
            self.input_file.set(filename)
            # Sugerir nome do arquivo de saída
            base_name = os.path.splitext(filename)[0]
            self.output_file.set(f"{base_name}_validado.csv")
            self.log(f"Arquivo de entrada selecionado: {filename}")
            
            # Verificar colunas do arquivo
            try:
                df = pd.read_csv(filename)
                self.log(f"Colunas encontradas: {', '.join(df.columns)}")
            except Exception as e:
                self.log(f"Erro ao ler o arquivo: {str(e)}")
    
    def browse_output(self):
        filename = filedialog.asksaveasfilename(
            title="Selecione o arquivo CSV de saída",
            filetypes=[("Arquivos CSV", "*.csv"), ("Todos os arquivos", "*.*")],
            defaultextension=".csv"
        )
        if filename:
            self.output_file.set(filename)
            self.log(f"Arquivo de saída definido: {filename}")
    
    def log(self, message):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert("end", f"[{timestamp}] {message}\n")
        self.log_text.see("end")
    
    def start_validation(self):
        if not self.input_file.get():
            messagebox.showerror("Erro", "Selecione um arquivo de entrada.")
            return
        
        if not self.output_file.get():
            messagebox.showerror("Erro", "Selecione um arquivo de saída.")
            return
        
        if self.processing:
            messagebox.showinfo("Informação", "Processo de validação já em andamento.")
            return
        
        # Iniciar processamento em uma thread separada
        self.processing = True
        self.process_button.configure(text="Processando...", state="disabled", fg_color="gray")
        thread = threading.Thread(target=self.validate_vats)
        thread.daemon = True
        thread.start()
    
    def validate_vats(self):
        try:
            input_file = self.input_file.get()
            output_file = self.output_file.get()
            vat_col = self.vat_column.get()
            country_col = self.country_column.get()
            
            self.log(f"Iniciando validação de VATs do arquivo {input_file}")
            
            # Ler o CSV
            df = pd.read_csv(input_file)
            
            # Verificar se as colunas existem
            if vat_col not in df.columns:
                self.log(f"Erro: Coluna {vat_col} não encontrada no arquivo")
                messagebox.showerror("Erro", f"Coluna {vat_col} não encontrada no arquivo")
                self.finish_processing()
                return
            
            if country_col not in df.columns:
                self.log(f"Aviso: Coluna {country_col} não encontrada. A validação será feita sem o código do país.")
            
            total_rows = len(df)
            self.log(f"Total de registros a processar: {total_rows}")
            
            # Adicionar colunas de resultado
            df['vat_valid'] = False
            df['vat_message'] = ''
            
            # Validar cada VAT
            for index, row in df.iterrows():
                vat_number = str(row[vat_col]).strip()
                country_code = str(row[country_col]).strip() if country_col in df.columns else None
                
                # Atualiza a barra de progresso
                progress = (index + 1) / total_rows
                self.update_progress(progress)
                
                try:
                    if not vat_number or vat_number.lower() == 'nan':
                        df.at[index, 'vat_message'] = 'VAT vazio'
                        continue
                    
                    max_retries = 3  # Número máximo de tentativas
                    retry_delay = 1  # Atraso entre tentativas em segundos
                    
                    for attempt in range(max_retries):
                        try:
                            # Se o VAT começa com GB, usa a VATLayer API
                            if vat_number.startswith("GB"):
                                api_key = "38eba532700f25cb9d6191ba121542aa"  # Substitua pela sua chave API
                                url = f"http://apilayer.net/api/validate?access_key={api_key}&vat_number={vat_number}"
                                response = requests.get(url)
                                response.raise_for_status()  # Lança uma exceção para códigos de status HTTP ruins
                                data = response.json()
                                
                                if data.get("valid"):
                                    df.at[index, 'vat_valid'] = True
                                    df.at[index, 'vat_message'] = 'Válido (VATLayer)'
                                else:
                                    df.at[index, 'vat_valid'] = False
                                    df.at[index, 'vat_message'] = f'Inválido (VATLayer): {data.get("error", {}).get("info", "Desconhecido")}'
                                break  # Sai do loop de repetição
                            else:
                                # Se temos o código do país, validamos com ele
                                if country_code and len(country_code) == 2:
                                    result = pyvat.check_vat_number(vat_number, country_code)
                                else:
                                    # Tentamos extrair o código do país do próprio VAT
                                    result = pyvat.check_vat_number(vat_number)
                                
                                df.at[index, 'vat_valid'] = result.is_valid
                                df.at[index, 'vat_message'] = 'Válido' if result.is_valid else 'Inválido'
                                break  # Se a validação for bem-sucedida, sai do loop de repetição
                        
                        except Exception as e:
                            if "MS_MAX_CONCURRENT_REQ" in str(e):
                                self.log(f"Erro MS_MAX_CONCURRENT_REQ ao validar VAT {vat_number}. Tentando novamente em {retry_delay} segundos (tentativa {attempt + 1}/{max_retries})")
                                time.sleep(retry_delay)  # Espera antes de tentar novamente
                            else:
                                df.at[index, 'vat_message'] = f"Erro: {str(e)}"
                                self.log(f"Erro ao validar VAT {vat_number}: {str(e)}")
                                break  # Se for um erro diferente, sai do loop de repetição
                    else:
                        # Se todas as tentativas falharem
                        df.at[index, 'vat_message'] = "Erro: Falha ao validar após várias tentativas"
                        self.log(f"Falha ao validar VAT {vat_number} após {max_retries} tentativas")
                    
                    if (index + 1) % 10 == 0 or index == 0:
                        self.log(f"Processados {index + 1} de {total_rows} VATs")
                
                except Exception as e:
                    df.at[index, 'vat_message'] = f"Erro: {str(e)}"
                    self.log(f"Erro ao validar VAT {vat_number}: {str(e)}")
            
            # Exportar resultados
            df.to_csv(output_file, index=False)
            self.log(f"Validação concluída. Resultados salvos em {output_file}")
            
            # Mostrar estatísticas
            valid_count = df['vat_valid'].sum()
            self.log(f"VATs válidos: {valid_count} de {total_rows} ({(valid_count/total_rows)*100:.2f}%)")
            
            messagebox.showinfo("Concluído", "Validação de VATs concluída com sucesso!")
        
        except Exception as e:
            self.log(f"Erro durante o processamento: {str(e)}")
            messagebox.showerror("Erro", f"Ocorreu um erro durante o processamento: {str(e)}")
        
        finally:
            self.finish_processing()
    
    def update_progress(self, value):
        self.progress_bar.set(value)
        self.update_idletasks()
    
    def finish_processing(self):
        self.processing = False
        self.process_button.configure(text="Validar VATs", state="normal", fg_color="green")
        self.progress_bar.set(0)

if __name__ == "__main__":
    app = VATValidatorApp()
    app.mainloop()