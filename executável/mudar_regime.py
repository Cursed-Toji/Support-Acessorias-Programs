from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import Select
from winotify import Notification
import time
import openpyxl
import customtkinter as tk
from tkinter import filedialog, messagebox

class AcessoriasBoot:
    def __init__(self, usuario, senha, caminho_planilha):
        self.usuario = usuario
        self.senha = senha
        self.caminho_planilha = caminho_planilha
        self.report_lines = []
        
        # Configuração do Firefox com opções
        self.options = webdriver.FirefoxOptions()
        '''self.options.add_argument('--headless')  # Executa em modo headless
        self.options.add_argument('--disable-gpu')  # Desabilita o GPU
        self.options.add_argument('window-size=945x1012')  # Tamanho da janela'''
        
        # Inicializa o WebDriver
        self.driver = webdriver.Firefox(options=self.options)

    def login(self):
        """Faz login no sistema Acessorias e navega até a aba de empresas."""
        driver = self.driver
        driver.get("https://app.acessorias.com")

        # Login
        campo_user = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//input[@name='mailAC']"))
        )
        campo_user.clear()
        campo_user.send_keys(self.usuario)
        time.sleep(3)

        campo_senha = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//input[@name='passAC']"))
        )
        campo_senha.clear()
        campo_senha.send_keys(self.senha)
        campo_senha.send_keys(Keys.RETURN)
        print('Fez login no Acessórias')

        # Notificação de login
        notificacao = Notification(app_id="honorarios.py", title="Login Realizado", msg="Login no Acessórias realizado com sucesso.")
        notificacao.show()

        # Aguarde alguns segundos para permitir que a página carregue após o login
        time.sleep(3)

        # Clica na aba empresa
        campo_empresa = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//a[@href='sysmain.php?m=4']"))
        )
        campo_empresa.click()
        time.sleep(3)

        # Processa os CNPJs
        self.processar_cnpjs(self.caminho_planilha)

    def carregar_cnpjs(self, caminho_planilha):
        """Carrega os CNPJs da planilha Excel."""
        wb = openpyxl.load_workbook(caminho_planilha)
        ws = wb.active
        cnpjs = []

        for row in ws.iter_rows(min_row=2, values_only=True):
            cnpj = row[0]  # A
            regime = row[1] # B
            cnpjs.append((cnpj, regime))  # Adiciona como tupla

        return cnpjs

    def processar_cnpjs(self, caminho_planilha):
        """Busca e interage com os CNPJs carregados da planilha."""
        driver = self.driver
        empresas = self.carregar_cnpjs(caminho_planilha)

        for cnpj, regime, in empresas:
            try:
                # Busca pelo CNPJ
                campo_busca = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, "//*[@id='searchString']"))
                )
                campo_busca.clear()
                campo_busca.send_keys(cnpj)
                campo_busca.send_keys(Keys.RETURN)
                time.sleep(2)
                print('Fez a busca pelo CNPJ', cnpj)

                # clica na empresa que realizou a busca
                campo_empresaId = driver.find_element(By.XPATH, "//*[@id='divEmpresas']/div[1]")
                campo_empresaId.click()
                time.sleep(2)

                # Clica no regime
                regime_alterar = driver.find_element(By.XPATH,"//*[@id='EmpRegID']")
                opcoes = Select(regime_alterar)
                opcoes.select_by_visible_text('Lucro Presumido')
                time.sleep(1)
                regime_alterar.send_keys(Keys.ENTER)
                time.sleep(2)

                # clica em ok para confirmar
                confirmar = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, "//*[@class='swal2-confirm btn btn-primary']"))
                )
                confirmar.click()
                time.sleep(2)

                # clica em salvar
                salvar_sair = driver.find_element(By.XPATH, "//*[@id='navFim']/div[2]/button[1]")
                salvar_sair.click()
                time.sleep(10)

                #clica em voltar
                voltar = driver.find_element(By.XPATH, "//*[@id='navFim']/div[2]/button[2]")
                voltar.click()
                time.sleep(4)
                
                notificacao = Notification(app_id="mudar_regime.py", title="Notificação Automação", msg=f"regime aplicado para o CNPJ: {cnpj}")
                notificacao.show()
                print('CNPJ com a nova obrigação', cnpj, '\n')
    
            except Exception as e:
                print(f"Erro ao aplicar o honorário {cnpj}: {e}")

        # Fecha o driver
        driver.quit()
    
    def save_report(self):
        """Salva o relatório de processamento em um arquivo de texto."""
        with open("report.txt", "w") as file:
            for line in self.report_lines:
                file.write(line + "\n")
        print("Relatório salvo como report.txt")

def iniciar_processo(usuario, senha, caminho_planilha):
    boot = AcessoriasBoot(usuario, senha, caminho_planilha)
    boot.login()
    boot.processar_cnpjs(caminho_planilha)
    boot.save_report()

def selecionar_planilha():
    caminho = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xlsx")])
    if caminho:
        caminho_planilha_var.set(caminho)

# Configuração da interface gráfica
root = tk.CTk()
root.title("Ferramenta Acessorias")

tk.CTkLabel(root, text="Usuário:").grid(row=0, column=0, padx=10, pady=5)
usuario_entry = tk.CTkEntry(root)
usuario_entry.grid(row=0, column=1, padx=10, pady=5)

tk.CTkLabel(root, text="Senha:").grid(row=1, column=0, padx=10, pady=5)
senha_entry = tk.CTkEntry(root, show='*')
senha_entry.grid(row=1, column=1, padx=10, pady=5)

tk.CTkLabel(root, text="Planilha:").grid(row=2, column=0, padx=10, pady=5)
caminho_planilha_var = tk.StringVar()
planilha_entry = tk.CTkEntry(root, textvariable=caminho_planilha_var, width=140)
planilha_entry.grid(row=2, column=1, padx=10, pady=5)

selecionar_planilha_button = tk.CTkButton(root, text="Selecionar", command=selecionar_planilha)
selecionar_planilha_button.grid(row=2, column=2, padx=10, pady=5)

def on_start():
    usuario = usuario_entry.get()
    senha = senha_entry.get()
    caminho_planilha = caminho_planilha_var.get()

    if not usuario or not senha or not caminho_planilha:
        messagebox.showerror("Erro", "Por favor, preencha todos os campos.")
    else:
        iniciar_processo(usuario, senha, caminho_planilha)

start_button = tk.CTkButton(root, text="Iniciar", command=on_start)
start_button.grid(row=3, column=0, columnspan=3, pady=10)

root.mainloop()