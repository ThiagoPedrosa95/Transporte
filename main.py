import tkinter as tk
from tkinter import messagebox, ttk
from threading import Thread
from datetime import datetime, timedelta
from playwright.sync_api import sync_playwright
from cryptography.fernet import Fernet
import time
import json
import os
import sys

# Detecta se o app está empacotado
if getattr(sys, 'frozen', False):
    BASE_DIR = getattr(sys, '_MEIPASS', os.path.dirname(__file__))
else:
    BASE_DIR = os.path.dirname(__file__)
    
# Caminhos dos arquivos para a mesma pasta do script
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CAMINHO_DADOS = os.path.join(SCRIPT_DIR, "dados_formulario.json")
CAMINHO_CREDENCIAIS = os.path.join(SCRIPT_DIR, "credentials.json")
CAMINHO_CHAVE = os.path.join(SCRIPT_DIR, "secret.key")

# Caminho onde o Chromium empacotado está
#chromium_path = os.path.join(BASE_DIR, "Lib/site-packages/playwright/driver/package/.local-browsers")
chromium_path = os.path.join("C:/Users/0360700/AppData/Local", "ms-playwright")
os.environ["PLAYWRIGHT_BROWSERS_PATH"] = chromium_path

# Utilidades para exportação do .ico do onibus
def caminho_icone(nome_arquivo):
    if getattr(sys, 'frozen', False):  # Executável via PyInstaller
        caminho_base = getattr(sys, '_MEIPASS', os.path.dirname(__file__))
    else:
        caminho_base = os.path.dirname(__file__)
    return os.path.join(caminho_base, nome_arquivo)

# Utilidades para criptografia de credenciais
def gerar_chave():
    """
    Gera uma chave de criptografia e a salva no caminho especificado.
    """
    chave = Fernet.generate_key()
    with open(CAMINHO_CHAVE, "wb") as f:
        f.write(chave)

def carregar_chave():
    """
    Carrega a chave de criptografia existente ou gera uma nova se não existir.

    Returns:
        bytes: Chave de criptografia.
    """
    if not os.path.exists(CAMINHO_CHAVE):
        gerar_chave()
    with open(CAMINHO_CHAVE, "rb") as f:
        return f.read()

def salvar_credenciais(login, senha):
    """
    Criptografa e salva o login e senha fornecidos.

    Args:
        login (str): Nome de usuário.
        senha (str): Senha do usuário.
    """
    chave = carregar_chave()
    f = Fernet(chave)
    dados = {
        "login": f.encrypt(login.encode()).decode(),
        "senha": f.encrypt(senha.encode()).decode()
    }
    with open(CAMINHO_CREDENCIAIS, "w") as f:
        json.dump(dados, f)

def carregar_credenciais():
    """
    Carrega e descriptografa as credenciais salvas, se existirem.

    Returns:
        tuple: login (str), senha (str)
    """
    if not os.path.exists(CAMINHO_CREDENCIAIS):
        return "", ""
    chave = carregar_chave()
    f = Fernet(chave)
    with open(CAMINHO_CREDENCIAIS, "r") as f_json:
        dados = json.load(f_json)
    try:
        login = f.decrypt(dados["login"].encode()).decode()
        senha = f.decrypt(dados["senha"].encode()).decode()
        return login, senha
    except:
        return "", ""

def salvar_dados(nome, endereco, roteiro):
    """
    Salva as informações de formulário fornecidas.

    Args:
        nome (str): Nome do usuário.
        endereco (str): Endereço.
        roteiro (str): Roteiro selecionado.
    """
    dados = {"nome": nome, "endereco": endereco, "roteiro": roteiro}
    with open(CAMINHO_DADOS, "w", encoding="utf-8") as f:
        json.dump(dados, f)

def carregar_dados():
    """
    Carrega os dados do formulário, se existirem.

    Returns:
        dict: Dados contendo nome, endereço e roteiro.
    """
    if os.path.exists(CAMINHO_DADOS):
        with open(CAMINHO_DADOS, "r", encoding="utf-8") as f:
            return json.load(f)
    return {"nome": "", "endereco": "", "roteiro": ""}

def obter_data_amanha():
    """
    Retorna a próxima data útil formatada como string.

    Returns:
        str: Data no formato "dd-mm-aaaa"
    """
    hoje = datetime.now()
    amanha = hoje + timedelta(days=1)
    if amanha.weekday() == 5:
        amanha += timedelta(days=2)
    elif amanha.weekday() == 6:
        amanha += timedelta(days=1)
    return amanha.strftime("%d-%m-%Y")

LOGIN = ""
PASSWORD = ""

ROTEIROS_MAP = {
    "01": "Roteiro 01 - Jacarepaguá",
    "02": "Roteiro 02 - Niterói",
    "03": "Roteiro 03 - Nova Iguaçu",
    "04": "Roteiro 04 - Petrópolis",
    "05": "Roteiro 05 - Realengo",
    "06": "Roteiro 06 - Recreio",
    "07": "Roteiro 07 - São Gonçalo",
    "08": "Roteiro 08 - Tijuca",
    "09": "Roteiro 09 - Jacarepaguá",
}

def launch_browser():
    """
    Inicia o navegador Chromium com Playwright.

    Returns:
        tuple: playwright, browser, page
    """
    playwright = sync_playwright().start()
    browser = playwright.chromium.launch(headless=False, args=["--start-maximized"])
    context = browser.new_context(viewport=None, ignore_https_errors=True)
    page = context.new_page()
    return playwright, browser, page

def acessar_url(page, url):
    """
    Acessa uma URL e aguarda carregamento completo.

    Args:
        page: Instância da página do Playwright.
        url (str): URL a ser acessada.
    """
    page.goto(url, timeout=60000)
    page.wait_for_load_state("networkidle")

def autenticar_home_page(page):
    """
    Autentica o usuário na página inicial do atendimento.

    Args:
        page: Página do navegador.
    """
    acessar_url(page, "https://atendimento.cepel.br")
    page.fill("#login_name", LOGIN)
    page.fill("#login_password", PASSWORD)
    page.wait_for_selector('button[name="submit"]', state="visible")
    page.click('button[name="submit"]')

def selecionar_radio_por_value(page, name: str, value: str, delay: float = 0.25):
    """
    Seleciona um botão de rádio pelo nome e valor.

    Args:
        page: Página Playwright.
        name (str): Nome do input radio.
        value (str): Valor a selecionar.
        delay (float): Tempo de espera após clicar.
    """
    seletor = f'input[type="radio"][name="{name}"][value="{value}"]'
    page.wait_for_selector(seletor, state="visible")
    page.click(seletor)
    time.sleep(delay)

def digitar_letra_por_letra(page, name: str, texto: str, delay: float = 0.05):
    """
    Digita um texto caractere por caractere no campo especificado.

    Args:
        page: Página do navegador.
        name (str): Nome do campo de input.
        texto (str): Texto a ser inserido.
        delay (float): Atraso entre letras.
    """
    seletor = f'input[name="{name}"]'
    page.click(seletor)
    for letra in texto:
        page.keyboard.insert_text(letra)
        time.sleep(delay)

def digitar_data(page, data: str, delay: float = 0.05):
    """
    Digita uma data caractere por caractere.

    Args:
        page: Página do navegador.
        data (str): Data no formato dd-mm-aaaa.
        delay (float): Atraso entre caracteres.
    """
    seletor = 'input[name="formcreator_field_893"] + input[type="text"]'
    page.click(seletor)
    for letra in data:
        page.keyboard.insert_text(letra)
        time.sleep(delay)
    

def selecao_roteiro(page, numero_roteiro: str):
    """
    Digita uma data caractere por caractere.

    Args:
        page: Página do navegador.
        numero_roteiro (str): Numero do roteiro de acordo com a lista.
    """
    numero_formatado = str(numero_roteiro).zfill(2)
    roteiro = ROTEIROS_MAP.get(numero_formatado)
    if not roteiro:
        raise ValueError(f"Roteiro '{numero_formatado}' não encontrado.")

    seletor = "span.select2-container--default[data-select2-id='10'] span.select2-selection__arrow"
    page.wait_for_selector(seletor, state="visible")
    page.click(seletor)
    time.sleep(0.5)
    page.wait_for_selector(f"//span[contains(text(), '{roteiro}')]")
    page.click(f"//span[contains(text(), '{roteiro}')]")

def enviar_pedido(page):
    """
    Envia o formulário preenchido.

    Args:
        page: Página do navegador.
    """
    button = 'button[type="submit"][value="Enviar"][name="add"]'
    page.wait_for_selector(button, state="visible")
    page.click(button)
    page.wait_for_load_state("networkidle")
    time.sleep(1)  # Aguarda a resposta do servidor

def criar_form_transporte(page, nome, local, numero_roteiro):
    """
    Acessa a página de formulário de transporte e preenche os campos necessários.

    Args:
        page: Instância da página Playwright.
        nome (str): Nome do usuário.
        local (str): Local de embarque.
        numero_roteiro (str): Código do roteiro (ex: "01").
    """
    acessar_url(page, "https://atendimento.cepel.br/marketplace/formcreator/front/formlist.php")
    page.click("//a[contains(text(), 'TRANSPORTE')]")
    page.wait_for_load_state("networkidle")

    RadioButtons = [
        {"name": "formcreator_field_700", "value": "Sim"},
        {"name": "formcreator_field_702", "value": "Roteiro"},
        {"name": "formcreator_field_703", "value": "Solicitar embarque"},
        {"name": "formcreator_field_705", "value": "Usuário interno"},
        {"name": "formcreator_field_707", "value": "Fundão"},
        {"name": "formcreator_field_706", "value": "Sim"},
        {"name": "formcreator_field_895", "value": "Não"},
    ]

    for item in RadioButtons:
        selecionar_radio_por_value(page, item["name"], item["value"])

    digitar_letra_por_letra(page, "formcreator_field_892", nome)
    digitar_letra_por_letra(page, "formcreator_field_894", local)
    digitar_data(page, obter_data_amanha())
    selecao_roteiro(page, numero_roteiro)
    enviar_pedido(page)


def executar_script(nome, local, numero_roteiro):
    """
    Executa o processo de automação de preenchimento do formulário de transporte.

    Args:
        nome (str): Nome do usuário.
        local (str): Endereço/local de embarque.
        numero_roteiro (str): Código do roteiro selecionado.
    """
    playwright, browser, page = launch_browser()
    try:
        autenticar_home_page(page)
        criar_form_transporte(page, nome, local, numero_roteiro)
        messagebox.showinfo("Concluído", "Tudo certo, pode clicar em enviar.")
    except Exception as e:
        messagebox.showerror("Erro", str(e))
    finally:
        browser.close()
        playwright.stop()


def abrir_janela_principal():
    """
    Cria a janela principal do formulário após o login, permitindo inserir os dados e executar o script.
    """
    main = tk.Toplevel()
    main.title("Formulário de Transporte")
    main.geometry("400x250")

    dados = carregar_dados()

    tk.Label(main, text="Nome:").pack(pady=5)
    entry_nome = tk.Entry(main, width=50)
    entry_nome.insert(0, dados.get("nome", ""))
    entry_nome.pack()

    tk.Label(main, text="Endereço:").pack(pady=5)
    entry_endereco = tk.Entry(main, width=50)
    entry_endereco.insert(0, dados.get("endereco", ""))
    entry_endereco.pack()

    tk.Label(main, text="Roteiro:").pack(pady=5)
    roteiro_combo = ttk.Combobox(main, width=47, state="readonly")
    roteiro_combo["values"] = list(ROTEIROS_MAP.values())
    roteiro_combo.set(dados.get("roteiro", ""))
    roteiro_combo.pack()

    def limpar_campos():
        """Limpa os campos de entrada do formulário."""
        entry_nome.delete(0, tk.END)
        entry_endereco.delete(0, tk.END)
        roteiro_combo.set("")

    def iniciar_execucao():
        """Valida os dados e inicia a execução do script em uma thread."""
        nome = entry_nome.get()
        local = entry_endereco.get()
        roteiro_selecionado = roteiro_combo.get()

        if not nome or not local or not roteiro_selecionado:
            messagebox.showwarning("Campos obrigatórios", "Preencha todos os campos.")
            return

        salvar_dados(nome, local, roteiro_selecionado)

        numero_roteiro = next((k for k, v in ROTEIROS_MAP.items() if v == roteiro_selecionado), None)
        if not numero_roteiro:
            messagebox.showerror("Erro", "Roteiro inválido selecionado.")
            return

        Thread(target=executar_script, args=(nome, local, numero_roteiro), daemon=True).start()

    frame_botoes = tk.Frame(main)
    frame_botoes.pack(pady=10)

    tk.Button(frame_botoes, text="Executar", command=iniciar_execucao, bg="green", fg="white").grid(row=0, column=0, padx=5)
    tk.Button(frame_botoes, text="Limpar Campos", command=limpar_campos).grid(row=0, column=1, padx=5)
    tk.Button(frame_botoes, text="Sair", command=lambda: (main.destroy(), root.quit()), bg="red", fg="white").grid(row=0, column=2, padx=5)
    

def tela_login():
    """
    Exibe a tela de login, validando credenciais e oferecendo a opção de salvamento.
    """
    global root
    root = tk.Tk()
    root.iconbitmap(caminho_icone("bus.ico"))
    root.title("Login")
    root.geometry("300x250")

    tk.Label(root, text="Login:").pack(pady=5)
    global entry_login
    entry_login = tk.Entry(root, width=40)
    entry_login.pack()

    tk.Label(root, text="Senha:").pack(pady=5)
    global entry_senha
    entry_senha = tk.Entry(root, show="*", width=40)
    entry_senha.pack()

    global var_salvar
    var_salvar = tk.BooleanVar()
    check_salvar = tk.Checkbutton(root, text="Salvar login e senha", variable=var_salvar)
    check_salvar.pack(pady=5)

    # Pré-preenche com credenciais salvas, se houver
    login_salvo, senha_salva = carregar_credenciais()
    if login_salvo and senha_salva:
        entry_login.insert(0, login_salvo)
        entry_senha.insert(0, senha_salva)
        var_salvar.set(True)

    def autenticar():
        """
        Verifica os campos de login, salva se necessário e abre a janela principal.
        """
        global LOGIN, PASSWORD
        login = entry_login.get()
        senha = entry_senha.get()

        if not login or not senha:
            messagebox.showwarning("Aviso", "Informe login e senha.")
            return

        LOGIN = login
        PASSWORD = senha

        if var_salvar.get():
            salvar_credenciais(login, senha)

        root.withdraw()
        abrir_janela_principal()

    frame_botoes = tk.Frame(root)
    frame_botoes.pack(pady=10)

    tk.Button(frame_botoes, text="Entrar", command=autenticar, bg="green", fg="white").grid(row=0, column=0, padx=5)
    tk.Button(frame_botoes, text="Limpar Campos", command=lambda: (entry_login.delete(0, tk.END), entry_senha.delete(0, tk.END))).grid(row=0, column=1, padx=5)
    tk.Button(frame_botoes, text="Sair", command=lambda: (root.destroy(), root.quit()), bg="red", fg="white").grid(row=0, column=2, padx=5)

    root.mainloop()


if __name__ == "__main__":
    tela_login()

# Compilar em um único executável
# pyinstaller --onefile --noconsole --name Agendador_Roteiro --hidden-import=playwright.async_api --add-data "C:\Users\0360700\AppData\Local\ms-playwright;ms-playwright" --add-data "secret.key;." --add-data "credentials.json;." --add-data "dados_formulario.json;." --add-data "bus.ico:." --icon=bus.ico interface_transporte.py