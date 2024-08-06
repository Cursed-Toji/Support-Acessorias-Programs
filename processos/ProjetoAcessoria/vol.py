from selenium import webdriver
from selenium.webdriver.support.select import Select 
from selenium.webdriver import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import time
import openpyxl

class AcessoriasBoot:
    def __init__(self, usuario, senha):
        self.usuario = usuario
        self.senha = senha
        self.driver = webdriver.Firefox()

#inicia o login no site da acessorias
    def login(self):
        driver = self.driver
        driver.get("https://app.acessorias.com")

        #login
        campo_user = driver.find_element(By.XPATH, "//input[@name='mailAC']")
        campo_user.clear()
        campo_user.send_keys(self.usuario)
        time.sleep(2)

        campo_senha = driver.find_element(By.XPATH, "//input[@name='passAC']")
        campo_senha.clear()
        campo_senha.send_keys(self.senha)
        campo_senha.send_keys(Keys.RETURN)

        # Aguarde alguns segundos para permitir que a página carregue após o login
        time.sleep(2)

        eocnitnuo = driver.find_element(By.XPATH, "//*[@id='M1']/a/b")
        ActionChains(driver)\
        .move_to_element(eocnitnuo)\
        .perform()
        time.sleep(2)

        econtinuo1 = driver.find_element(By.XPATH, "//*[@id='M130']/a")
        ActionChains(driver)\
        .move_to_element(econtinuo1)\
        .perform()
        time.sleep(2)

        econtinuo2 = driver.find_element(By.XPATH, "//*[@id='M128']")
        ActionChains(driver)\
        .move_to_element(econtinuo2)\
        .perform()
        econtinuo2.click()
        time.sleep(3)

       # butao = driver.find_element(By.XPATH, "//*[@id='btShowFilters']")
       # butao.click()


        '''
        burao2 = "Com alguma obrigação correspondente"

        achados = driver.find_element(By.XPATH,"//*[@id='fieldFilters']/span/span[1]/span/ul/li/input")
        achados.send_keys(butao1)
        achados.send_keys(Keys.RETURN)
        achados.send_keys(burao2)
        achados.send_keys(Keys.RETURN)
        time.sleep(4)

        clicarFiltrar = driver.find_element(By.XPATH, "//button[@id='btFilter']")
        clicarFiltrar.click()
        time.sleep(9000)
        '''
        editar = driver.find_element(By.XPATH, "/html/body/div[2]/div[2]/div[2]/div/div/div[3]/div[1]/span/a[2]")
        editar.click()
        time.sleep(2)

        agendado = "Sim - Agendado"
        emailimediato = driver.find_element(By.XPATH,"//*[@id='EclSendMail']")
        emailimediato.send_keys(agendado)
        emailimediato.send_keys(Keys.RETURN)
        time.sleep(3)

        botaoSalvar = driver.find_element(By.XPATH, '/html/body/div[2]/div[2]/div[2]/div/div/form/div[4]/div[3]/div/button[1]')
        botaoSalvar.click()






     
        
# Criar uma instância da classe AcessoriasBoot
boot = AcessoriasBoot('cenilda@escritorioscheffer.com.br', 'teste123')
#boot = AcessoriasBoot('Usuario', 'Senha')
boot.login()