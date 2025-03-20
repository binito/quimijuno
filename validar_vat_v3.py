import customtkinter as ctk
from tkinter import filedialog, messagebox
import os
import pandas as pd
import requests
import re
import threading
import time
from typing import Dict, List, Tuple, Union
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, WebDriverException
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException, ElementNotInteractableException

class VATValidator:
    """
    Classe para validar números de VAT de diferentes países
    usando várias APIs e métodos de validação.
    """
    
    def __init__(self):
        self.vies_countries = [
            'AT', 'BE', 'BG', 'CY', 'CZ', 'DE', 'DK', 'EE', 'ES', 'FI', 'FR', 
            'GR', 'EL', 'HR', 'HU', 'IE', 'IT', 'LT', 'LU', 'LV', 'MT', 
            'NL', 'PL', 'PT', 'RO', 'SE', 'SI', 'SK'
        ]  # Removed GB from VIES countries
        self.special_validators = {
            'CH': self.validate_ch_vat,
            'US': self.validate_us_vat,
            'CA': self.validate_ca_vat,
            'AR': self.validate_ar_vat,
            'CN': self.validate_cn_vat,
            'SA': self.validate_sa_vat,
            'ZA': self.validate_za_vat,
            'TK': self.validate_tk_vat,
            'GB': self.validate_gb_vat,  # Added GB to special validators
        }
        self.driver = None
        self.setup_selenium()
    
    def setup_selenium(self):
        """Configura o browser Selenium para validação GB."""
        try:
            options = webdriver.ChromeOptions()
            options.add_argument('--headless')
            options.add_argument('--no-sandbox')
            options.add_argument('--disable-dev-shm-usage')
            options.add_argument('--disable-gpu')
            options.add_argument('--window-size=1920,1080')
            options.add_argument('--ignore-certificate-errors')
            options.add_argument('--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36')
            
            service = Service(ChromeDriverManager().install())
            self.driver = webdriver.Chrome(service=service, options=options)
            self.driver.set_page_load_timeout(30)
            print("Selenium setup successful")
        except Exception as e:
            print(f"Erro ao configurar Selenium: {str(e)}")
            self.driver = None
    
    def load_vat_data(self, file_path: str) -> pd.DataFrame:
        """
        Carrega os dados de VAT de um arquivo CSV.
        
        Args:
            file_path: Caminho para o arquivo CSV
            
        Returns:
            DataFrame com os números de VAT
        """
        try:
            # Tenta carregar como CSV padrão
            df = pd.read_csv(file_path)
            # Verifica se tem coluna 'vat'
            if 'vat' not in df.columns:
                # Se não tiver, tenta outras abordagens
                df = pd.read_csv(file_path, header=None)
                df.columns = ['vat']
        except Exception as e:
            # Se falhar, tenta ler o arquivo como texto
            with open(file_path, 'r') as file:
                lines = [line.strip() for line in file.readlines()]
                df = pd.DataFrame(lines, columns=['vat'])
        
        # Remover espaços em branco nos números de VAT
        df['vat'] = df['vat'].str.strip()
        return df
    
    def extract_country_code(self, vat: str) -> str:
        """
        Extrai o código do país de um número de VAT.
        
        Args:
            vat: Número de VAT
            
        Returns:
            Código do país
        """
        # Padrões comuns para códigos de países (2 primeiros caracteres alfabéticos)
        country_match = re.match(r'^([A-Z]{2})', vat)
        if country_match:
            return country_match.group(1)
        
        # Casos especiais
        if vat.startswith('EL'):  # Grécia (usado em alguns contextos ao invés de GR)
            return 'EL'
        
        # Alguns casos especiais onde o código do país está após outros caracteres
        special_patterns = {
            r'U\.A\.E\.': 'AE',
            r'CHE\s*': 'CH',
            r'DB': 'DB',
            r'DMCC': 'AE',  # Dubai Multi Commodities Centre
            r'IEC': 'IN',  # Import Export Code da Índia
            r'JPY': 'JP',
            r'KT': 'KW',  # Kuwait
            r'TN': 'TN',  # Tunísia
            r'CS': 'CS',  # Sérvia e Montenegro (código antigo)
        }
        
        for pattern, code in special_patterns.items():
            if re.match(pattern, vat):
                return code
        
        # Se não conseguir identificar, retorna os dois primeiros caracteres
        if len(vat) >= 2 and vat[:2].isalpha():
            return vat[:2]
        
        # Caso não consiga identificar o país
        return "UNKNOWN"
    
    def validate_vat(self, vat_list: List[str], progress_callback=None) -> Dict[str, Dict[str, Union[bool, str]]]:
        """
        Valida uma lista de números de VAT.
        
        Args:
            vat_list: Lista de números de VAT para validar
            progress_callback: Função de callback para atualizar o progresso
            
        Returns:
            Dicionário com os resultados da validação
        """
        results = {}
        total = len(vat_list)
        
        for i, vat in enumerate(vat_list):
            # Atualiza o progresso se houver um callback
            if progress_callback:
                progress_callback(i+1, total)
            
            # Inicializa o resultado
            results[vat] = {
                'valid': False,
                'method': 'unknown',
                'message': 'Não validado'
            }
            
            # Remove espaços em branco
            vat_clean = vat.strip()
            
            if not vat_clean:
                results[vat]['message'] = 'Número de VAT vazio'
                continue
            
            # Extrai o código do país
            country_code = self.extract_country_code(vat_clean)
            
            # Valida de acordo com o país
            if country_code in self.vies_countries:  # GB will now use VIES
                valid, message = self.validate_vies_vat(vat_clean, country_code)
                results[vat] = {
                    'valid': valid,
                    'method': 'VIES',
                    'message': message
                }
            elif country_code in self.special_validators:
                valid, message = self.special_validators[country_code](vat_clean)
                results[vat] = {
                    'valid': valid,
                    'method': f'Special ({country_code})',
                    'message': message
                }
            else:
                results[vat] = {
                    'valid': False,
                    'method': 'Unknown',
                    'message': f'Método de validação não disponível para o país {country_code}'
                }
        
        return results
    
    def validate_vies_vat(self, vat: str, country_code: str) -> Tuple[bool, str]:
        """
        Valida um número de VAT usando a API REST do VIES da União Europeia.
        
        Este método foi adaptado para utilizar o método presente no ficheiro "validar_vat.py".
        
        Args:
            vat: Número de VAT
            country_code: Código do país
            
        Returns:
            Tupla (válido, mensagem)
        """
        try:
            # Remove o código do país do VAT se já estiver presente
            if vat.startswith(country_code):
                vat_number = vat[len(country_code):]
            else:
                vat_number = vat
            
            # Monta a URL da API REST do VIES
            url = f"https://ec.europa.eu/taxation_customs/vies/rest-api/ms/{country_code}/vat/{vat_number}"
            headers = {"Accept": "application/json"}
            
            # Realiza a requisição GET à API do VIES
            response = requests.get(url, headers=headers, timeout=10)
            
            # Verifica se a resposta foi bem-sucedida
            if response.status_code == 200:
                data = response.json()
                if data.get("isValid", False):
                    return True, "Número de VAT válido (VIES)"
                else:
                    return False, "Número de VAT inválido (VIES)"
            else:
                return False, "Erro na API VIES"
        except Exception as e:
            # Em caso de exceção, retorna uma mensagem de erro de conexão
            return False, "Erro de conexão"
    
    def validate_vat_format(self, vat: str, country_code: str) -> Tuple[bool, str]:
        """
        Valida o formato do VAT de acordo com regras específicas de cada país.
        Esta é uma validação secundária quando a API VIES não está disponível.
        
        Args:
            vat: Número de VAT
            country_code: Código do país
            
        Returns:
            Tupla (válido, mensagem)
        """
        # Remove o código do país se já estiver no início do VAT
        if vat.startswith(country_code):
            vat_number = vat[len(country_code):]
        else:
            vat_number = vat
        
        # Regras de validação de formato por país
        format_rules = {
            'PT': r'^\d{9}$',  # Portugal: 9 dígitos
            'ES': r'^[A-Z0-9]\d{7}[A-Z0-9]$',  # Espanha: letra/número + 7 dígitos + letra/número
            'FR': r'^\d{11}$|^[A-Z]{2}\d{9}$',  # França: 11 dígitos ou 2 letras + 9 dígitos
            'DE': r'^\d{9}$',  # Alemanha: 9 dígitos
            'IT': r'^\d{11}$',  # Itália: 11 dígitos
            'GB': r'^\d{9}$|^\d{12}$|^GD\d{3}$|^HA\d{3}$',  # Reino Unido: vários formatos
        }
        
        # Se existe uma regra para o país, valida o formato
        if country_code in format_rules:
            pattern = format_rules[country_code]
            if re.match(pattern, vat_number):
                return True, f"Formato do VAT válido para {country_code}"
            else:
                return False, f"Formato do VAT inválido para {country_code}"
        
        # Caso não haja regra específica, considera válido pelo formato
        return True, f"Formato não verificado para {country_code}"
    
    def validate_ch_vat(self, vat: str) -> Tuple[bool, str]:
        """Valida um número de VAT suíço."""
        vat_clean = re.sub(r'^CH[E]?\s*', '', vat)
        pattern = r'^\d{9}$|^\d{3}\.\d{3}\.\d{3}$'
        if re.match(pattern, vat_clean):
            return True, "Formato de VAT suíço válido"
        return False, "Formato de VAT suíço inválido"
    
    def validate_us_vat(self, vat: str) -> Tuple[bool, str]:
        """Valida um EIN (Employer Identification Number) dos EUA."""
        vat_clean = re.sub(r'^US\s*', '', vat)
        pattern = r'^\d{2}-\d{7}$|^\d{9}$|^\d{2}\d{7}$'
        if re.match(pattern, vat_clean):
            return True, "Formato de EIN americano válido"
        return False, "Formato de EIN americano inválido"
    
    def validate_ca_vat(self, vat: str) -> Tuple[bool, str]:
        """Valida um número de BN (Business Number) canadense."""
        vat_clean = re.sub(r'^CA\s*', '', vat)
        pattern = r'^\d{9}$|^\d{9}RT\d{4}$'
        if re.match(pattern, vat_clean):
            return True, "Formato de BN canadense válido"
        return False, "Formato de BN canadense inválido"
    
    def validate_ar_vat(self, vat: str) -> Tuple[bool, str]:
        """Valida um CUIT (Código Único de Identificação Tributária) argentino."""
        vat_clean = re.sub(r'^AR\s*', '', vat)
        pattern = r'^\d{11}$|^\d{2}-\d{8}-\d{1}$'
        if re.match(pattern, vat_clean):
            return True, "Formato de CUIT argentino válido"
        return False, "Formato de CUIT argentino inválido"
    
    def validate_cn_vat(self, vat: str) -> Tuple[bool, str]:
        """Validação para TIN chinês."""
        vat_clean = re.sub(r'^CN\s*', '', vat)
        if len(vat_clean) >= 15 and len(vat_clean) <= 20:
            return True, "Formato de TIN chinês potencialmente válido"
        return False, "Formato de TIN chinês inválido"
    
    def validate_sa_vat(self, vat: str) -> Tuple[bool, str]:
        """Validação para TIN da Arábia Saudita."""
        vat_clean = re.sub(r'^SA\s*', '', vat)
        if re.match(r'^\d{15}$', vat_clean):
            return True, "Formato de TIN saudita válido"
        return False, "Formato de TIN saudita inválido"
    
    def validate_za_vat(self, vat: str) -> Tuple[bool, str]:
        """Validação para VAT da África do Sul."""
        vat_clean = re.sub(r'^ZA\s*', '', vat)
        if re.match(r'^\d{10}$', vat_clean):
            return True, "Formato de VAT sul-africano válido"
        return False, "Formato de VAT sul-africano inválido"
    
    def validate_tk_vat(self, vat: str) -> Tuple[bool, str]:
        """Validação para VAT da Turquia."""
        vat_clean = re.sub(r'^TK\s*', '', vat)
        if re.match(r'^\d{10}$', vat_clean):
            return True, "Formato de VAT turco válido"
        return False, "Formato de VAT turco inválido"
    
    def validate_gb_vat(self, vat: str) -> Tuple[bool, str]:
        """Valida um número de VAT do Reino Unido usando web scraping do site oficial."""
        vat_clean = re.sub(r'^GB\s*', '', vat).strip()
        
        if not self.driver:
            print("Selenium not available, using format validation")
            return self.validate_vat_format(vat_clean, "GB")
        
        try:
            # Acessa a página de validação
            print(f"Validating GB VAT: {vat_clean}")
            self.driver.get("https://www.tax.service.gov.uk/check-vat-number/enter-vat-details")
            
            try:
                # Tenta aceitar cookies se o diálogo estiver presente
                cookie_button = WebDriverWait(self.driver, 5).until(
                    EC.element_to_be_clickable((By.ID, "accept-cookies"))
                )
                cookie_button.click()
                print("Accepted cookies")
            except:
                print("No cookie consent needed")
            
            # Aguarda e preenche o campo VAT
            try:
                vat_input = WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located((By.ID, "vatNumber"))
                )
                vat_input.clear()
                vat_input.send_keys(vat_clean)
                vat_input.send_keys(Keys.TAB)  # Move to next field
                print("VAT number entered")
                
                # Clica no botão usando JavaScript
                continue_button = self.driver.find_element(By.ID, "continue")
                self.driver.execute_script("arguments[0].click();", continue_button)
                print("Continue button clicked")
                
                # Aguarda resultado
                try:
                    # Primeiro tenta encontrar mensagem de sucesso
                    result_element = WebDriverWait(self.driver, 10).until(
                        EC.presence_of_element_located((By.CLASS_NAME, "govuk-panel__title"))
                    )
                    print(f"Result found: {result_element.text}")
                    if "valid" in result_element.text.lower():
                        return True, "Número de VAT GB válido (HMRC website)"
                except:
                    # Se não encontrar sucesso, procura erro
                    try:
                        error_element = WebDriverWait(self.driver, 5).until(
                            EC.presence_of_element_located((By.CLASS_NAME, "govuk-error-summary"))
                        )
                        print("Error message found")
                        return False, "Número de VAT GB inválido (HMRC website)"
                    except:
                        print("No result found")
            except Exception as e:
                print(f"Error during validation: {str(e)}")
                raise
                
        except Exception as e:
            print(f"GB VAT validation error: {str(e)}")
            try:
                self.driver.quit()
                self.setup_selenium()
            except:
                pass
            
            # Fallback to format validation
            format_valid, _ = self.validate_vat_format(vat_clean, "GB")
            if format_valid:
                return True, "Número de VAT GB com formato válido (erro no site HMRC)"
            return False, "Número de VAT GB inválido"
        
        return False, "Falha na validação do VAT GB"
    
    def __del__(self):
        """Cleanup do Selenium quando o objeto é destruído."""
        if self.driver:
            try:
                self.driver.quit()
            except:
                pass
    
    def save_validation_results(self, results: Dict[str, Dict[str, Union[bool, str]]], output_file: str) -> None:
        """
        Salva os resultados da validação em um arquivo CSV.
        
        Args:
            results: Resultados da validação
            output_file: Nome do arquivo de saída
        """
        # Converte o dicionário de resultados para DataFrame
        data = []
        for vat, result in results.items():
            data.append({
                'vat': vat,
                'valid': result['valid'],
                'method': result['method'],
                'message': result['message']
            })
        
        df = pd.DataFrame(data)
        df.to_csv(output_file, index=False)
        return len(data), sum(1 for result in results.values() if result['valid'])


class VATValidatorApp:
    def __init__(self, root):
        self.root = root
        # Configura o tamanho da janela principal
        self.root.title("Validador de VAT")
        self.root.geometry("900x650")
        
        self.validator = VATValidator()
        self.input_file = ""
        self.output_file = ""
        
        # Flag para cancelamento da validação
        self.cancel_requested = False
        
        # Variáveis de controle
        self.status_var = ctk.StringVar(value="Pronto para começar")
        self.input_file_var = ctk.StringVar(value="Nenhum arquivo selecionado")
        self.output_file_var = ctk.StringVar(value="resultados_validacao.csv")
        
        # Criar interface
        self.create_widgets()
        
        # Centralizar janela
        self.center_window()
    
    def center_window(self):
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f"{width}x{height}+{x}+{y}")
    
    def create_widgets(self):
        # Frame principal
        self.main_frame = ctk.CTkFrame(self.root)
        self.main_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # Título
        title_label = ctk.CTkLabel(self.main_frame, text="Validador de Números VAT", font=("Arial", 18, "bold"))
        title_label.pack(pady=(0,20))
        
        # Frame de seleção de arquivo
        file_frame = ctk.CTkFrame(self.main_frame)
        file_frame.pack(fill="x", pady=10)
        
        file_label = ctk.CTkLabel(file_frame, textvariable=self.input_file_var)
        file_label.pack(side="left", padx=5, fill="x", expand=True)
        
        browse_button = ctk.CTkButton(file_frame, text="Procurar...", command=self.browse_input_file)
        browse_button.pack(side="right", padx=5)
        
        # Frame de arquivo de saída
        output_frame = ctk.CTkFrame(self.main_frame)
        output_frame.pack(fill="x", pady=10)
        
        output_label = ctk.CTkLabel(output_frame, text="Nome do arquivo:")
        output_label.pack(side="left", padx=5)
        
        output_entry = ctk.CTkEntry(output_frame, textvariable=self.output_file_var)
        output_entry.pack(side="left", padx=5, fill="x", expand=True)
        
        # Frame para botões de validação e cancelamento
        button_frame = ctk.CTkFrame(self.main_frame)
        button_frame.pack(fill="x", pady=10)
        
        validate_button = ctk.CTkButton(button_frame, text="Validar VATs", command=self.start_validation)
        validate_button.pack(side="left", pady=10, padx=5)
        
        cancel_button = ctk.CTkButton(button_frame, text="Cancelar", command=self.cancel_validation)
        cancel_button.pack(side="left", pady=10, padx=5)
        
        # Barra de progresso
        progress_frame = ctk.CTkFrame(self.main_frame)
        progress_frame.pack(fill="x", pady=10)
        
        self.progress_bar = ctk.CTkProgressBar(progress_frame)
        self.progress_bar.set(0)
        self.progress_bar.pack(fill="x", padx=5, pady=5)
        
        # Status
        status_frame = ctk.CTkFrame(self.main_frame)
        status_frame.pack(fill="x", pady=10)
        
        status_label = ctk.CTkLabel(status_frame, textvariable=self.status_var)
        status_label.pack(side="left", padx=5)
        
        # Área para exibir os últimos resultados
        self.results_frame = ctk.CTkFrame(self.main_frame)
        self.results_frame.pack(fill="both", expand=True, pady=10)
        
        self.results_textbox = ctk.CTkTextbox(self.results_frame, wrap="word")
        self.results_textbox.pack(fill="both", expand=True, padx=5, pady=5)
    
    def browse_input_file(self):
        file_path = filedialog.askopenfilename(
            title="Selecione o arquivo CSV com VATs",
            filetypes=[("Arquivos CSV", "*.csv"), ("Todos os arquivos", "*.*")]
        )
        
        if file_path:
            self.input_file = file_path
            self.input_file_var.set(os.path.basename(file_path))
            base_name = os.path.splitext(os.path.basename(file_path))[0]
            self.output_file_var.set(f"{base_name}_validados.csv")
    
    def update_progress(self, current, total):
        progress = current / total
        self.progress_bar.set(progress)
        self.status_var.set(f"Validando {current} de {total} VATs...")
        self.root.update_idletasks()
    
    def append_log(self, message: str):
        self.results_textbox.insert("end", message)
        self.results_textbox.see("end")
    
    def start_validation(self):
        if not self.input_file:
            messagebox.showerror("Erro", "Por favor, selecione um arquivo de entrada.")
            return
        
        output_file = self.output_file_var.get()
        if not output_file.endswith('.csv'):
            output_file += '.csv'
        
        # Reinicia o flag de cancelamento e a área de resultados
        self.cancel_requested = False
        self.results_textbox.delete("1.0", "end")
        self.status_var.set("Iniciando validação...")
        self.progress_bar.set(0)
        self.results_textbox.insert("end", "Validando... por favor aguarde\n")
        
        validation_thread = threading.Thread(
            target=self.run_validation,
            args=(self.input_file, output_file)
        )
        validation_thread.daemon = True
        validation_thread.start()
    
    def cancel_validation(self):
        self.cancel_requested = True
        self.status_var.set("Cancelando validação...")
    
    def run_validation(self, input_file, output_file):
        try:
            df = self.validator.load_vat_data(input_file)
            vat_list = df['vat'].tolist()
            results = {}
            total = len(vat_list)
            
            # Loop para processar cada VAT e atualizar a interface em tempo real
            for i, vat in enumerate(vat_list):
                # Verifica se foi solicitado cancelamento
                if self.cancel_requested:
                    self.root.after(0, self.validation_cancelled)
                    return
                
                # Atualiza a barra de progresso
                self.root.after(0, self.update_progress, i+1, total)
                
                # Valida o VAT individualmente
                result = self.validator.validate_vat([vat])[vat]
                results[vat] = result
                
                # Prepara uma mensagem de resultado para exibição
                status_icon = "✓" if result['valid'] else "✗"
                msg = f"{status_icon} {vat} - {result['message']}\n"
                self.root.after(0, self.append_log, msg)
                
                # Pequena pausa para visualizar as atualizações (opcional)
                time.sleep(0.05)
            
            total_count, valid_count = self.validator.save_validation_results(results, output_file)
            self.root.after(0, self.show_results, output_file, total_count, valid_count)
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("Erro", f"Ocorreu um erro durante a validação: {str(e)}"))
            self.root.after(0, lambda: self.status_var.set("Erro na validação."))
    
    def validation_cancelled(self):
        self.status_var.set("Validação cancelada pelo usuário.")
        self.append_log("\nValidação cancelada.\n")
    
    def show_results(self, output_file, total, valid):
        self.results_textbox.delete("1.0", "end")
        result_text = (f"Validação concluída com sucesso!\n\n"
                       f"Total de VATs validados: {total}\n"
                       f"VATs válidos: {valid} ({valid/total*100:.2f}%)\n"
                       f"VATs inválidos: {total-valid} ({(total-valid)/total*100:.2f}%)\n\n"
                       f"Resultados salvos em: {output_file}")
        self.results_textbox.insert("end", result_text)
        self.status_var.set("Validação concluída!")
        self.progress_bar.set(1.0)
    
    def open_results_file(self, file_path):
        try:
            import subprocess
            import platform
            system = platform.system()
            if system == 'Windows':
                os.startfile(file_path)
            elif system == 'Darwin':
                subprocess.call(['open', file_path])
            else:
                subprocess.call(['xdg-open', file_path])
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível abrir o arquivo: {str(e)}")


# Iniciar a aplicação
if __name__ == "__main__":
    app_root = ctk.CTk()
    app_root.state("zoomed")  # Maximiza a janela
    app = VATValidatorApp(app_root)
    app_root.mainloop()
