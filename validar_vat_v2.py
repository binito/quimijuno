import pandas as pd
import requests
import customtkinter as ctk
import tkinter.filedialog as filedialog
import threading

# Variável global para sinalizar paragem
stop_flag = False

# Função para validar VATs de países não-UK utilizando a API VIES
def validar_vat_vies(vat):
    if len(vat) < 3:
        return "NIF inválido"
    country = vat[:2]
    vat_number = vat[2:]
    url = f"https://ec.europa.eu/taxation_customs/vies/rest-api/ms/{country}/vat/{vat_number}"
    headers = {"Accept": "application/json"}
    try:
        response = requests.get(url, headers=headers, timeout=5)
        if response.status_code == 200:
            data = response.json()
            return "Válido" if data.get("isValid", False) else "NIF inválido"
        else:
            return "Erro na API"
    except Exception as e:
        return f"Erro de conexão: {e}"

# Função para validar VAT do Reino Unido utilizando a API do vatlayer
def validar_vat_uk_thirdparty(vat):
    """
    Valida um número de VAT do Reino Unido usando a API do vatlayer.
    É necessário registar-se em https://vatlayer.com/ para obter uma chave (access_key).
    """
    vat_converted = "GB" + vat[2:]
    access_key = "YOUR_ACCESS_KEY"  # Substituir pela sua chave
    url = f"https://apilayer.net/api/validate?access_key={access_key}&vat_number={vat_converted}"
    try:
        response = requests.get(url, timeout=10)
        if response.status_code == 200:
            data = response.json()
            return "Válido" if data.get("valid", False) else "NIF inválido"
        else:
            return "Erro na API"
    except Exception as e:
        return f"Erro de conexão: {e}"

# Função principal que escolhe o método de validação
def validar_vat(vat):
    if vat.startswith("UK"):
        return validar_vat_uk_thirdparty(str(vat))
    else:
        return validar_vat_vies(str(vat))

# Função chamada pelo botão "Iniciar Validação"
def iniciar_validacao():
    global stop_flag
    stop_flag = False  # Sempre que iniciamos uma validação, o flag é reposto

    origem = entry_origem.get()
    destino = entry_destino.get()
    
    if not origem:
        label_status.configure(text="Por favor, selecione o ficheiro de origem.")
        return
    if not destino:
        label_status.configure(text="Por favor, selecione o ficheiro de destino.")
        return

    try:
        df = pd.read_excel(origem)
    except Exception as e:
        label_status.configure(text=f"Erro ao carregar o ficheiro: {e}")
        return

    total = len(df)
    if total == 0:
        label_status.configure(text="O ficheiro não contém dados.")
        return

    def processar():
        for index, row in df.iterrows():
            # Verifica se foi pedido para parar
            if stop_flag:
                # Se pararmos, terminamos o loop
                break

            vat = row.get("vat")
            resultado = validar_vat(str(vat))
            df.at[index, "VAT_Validado"] = resultado

            progresso = (index + 1) / total
            progress_bar.set(progresso)
            label_progresso.configure(text=f"Progresso: {int(progresso * 100)}%")

            log_textbox.insert("end", f"NIF: {vat} - {resultado}\n")
            log_textbox.see("end")
            
            root.update_idletasks()

        # Se parou antes de concluir, mensagem de interrupção
        if stop_flag:
            label_status.configure(text="Validação interrompida.")
        else:
            # Caso contrário, grava o resultado final
            try:
                df.to_excel(destino, index=False)
                label_status.configure(text="Validação concluída com sucesso!")
            except Exception as e:
                label_status.configure(text=f"Erro ao guardar o ficheiro: {e}")

    # Inicia a thread
    thread = threading.Thread(target=processar)
    thread.start()

# Função chamada pelo botão "Parar Validação"
def parar_validacao():
    global stop_flag
    stop_flag = True  # Ativa o sinal para parar o processo

# Função para selecionar o ficheiro de origem
def selecionar_ficheiro_origem():
    ficheiro = filedialog.askopenfilename(
        title="Selecione o ficheiro de origem",
        filetypes=[("Ficheiros Excel", "*.xlsx *.xls")]
    )
    if ficheiro:
        entry_origem.delete(0, ctk.END)
        entry_origem.insert(0, ficheiro)

# Função para selecionar o ficheiro de destino
def selecionar_ficheiro_destino():
    ficheiro = filedialog.asksaveasfilename(
        title="Selecione o ficheiro de destino",
        defaultextension=".xlsx",
        filetypes=[("Ficheiros Excel", "*.xlsx")]
    )
    if ficheiro:
        entry_destino.delete(0, ctk.END)
        entry_destino.insert(0, ficheiro)

# Criação da janela principal
root = ctk.CTk()
root.title("Validação de VATs")
root.geometry("600x580")

# Ficheiro de origem
label_origem = ctk.CTkLabel(root, text="Ficheiro de Origem:")
label_origem.pack(pady=(20, 5))
entry_origem = ctk.CTkEntry(root, width=400)
entry_origem.pack(pady=(0, 5))
botao_origem = ctk.CTkButton(root, text="Selecionar Ficheiro", command=selecionar_ficheiro_origem)
botao_origem.pack(pady=(0, 20))

# Ficheiro de destino
label_destino = ctk.CTkLabel(root, text="Ficheiro de Destino:")
label_destino.pack(pady=(0, 5))
entry_destino = ctk.CTkEntry(root, width=400)
entry_destino.pack(pady=(0, 5))
botao_destino = ctk.CTkButton(root, text="Selecionar Ficheiro", command=selecionar_ficheiro_destino)
botao_destino.pack(pady=(0, 20))

# Botão para iniciar
botao_validar = ctk.CTkButton(root, text="Iniciar Validação", command=iniciar_validacao)
botao_validar.pack(pady=(0, 10))

# Botão para parar
botao_parar = ctk.CTkButton(root, text="Parar Validação", command=parar_validacao)
botao_parar.pack(pady=(0, 20))

# Barra de progresso
progress_bar = ctk.CTkProgressBar(root, width=400)
progress_bar.set(0)
progress_bar.pack(pady=(0, 10))

# Rótulo de progresso
label_progresso = ctk.CTkLabel(root, text="Progresso: 0%")
label_progresso.pack()

# Caixa de texto para log
log_textbox = ctk.CTkTextbox(root, width=500, height=150)
log_textbox.pack(pady=(10, 10))

# Rótulo de status
label_status = ctk.CTkLabel(root, text="")
label_status.pack(pady=(10, 0))

root.mainloop()
