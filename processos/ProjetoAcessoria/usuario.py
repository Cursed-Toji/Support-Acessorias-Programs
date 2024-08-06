from selenium import webdriver
from selenium.webdriver.support.select import Select 
from selenium.webdriver import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import time
import main

class usuarioAcessorias:
    def __init__(self, driver):
        self.usuario = main.set_temp_text
        self.senha = main.clear_temp_text
        self.driver = driver

    def fazerLogin(self):
        driver = self.driver
        campo_user = driver.find_element(By.XPATH, "//input[@name='mailAC']")
        campo_user.clear()
        campo_user.send_keys(self.usuario)
        time.sleep(2)

        campo_senha = driver.find_element(By.XPATH, "//input[@name='passAC']")
        campo_senha.clear()
        campo_senha.send_keys(self.senha)
        campo_senha.send_keys(Keys.RETURN)
        time.sleep(9000)

    def marcarDepartamento(self):
        driver = self.driver
        econtinuo = driver.find_element(By.XPATH, "//*[@id='M1']/a/b")
        ActionChains(driver)\
        .move_to_element(econtinuo)\
        .perform()
        time.sleep(2)

        econtinuoUm = driver.find_element(By.XPATH, "//*[@id='M130']/a")
        ActionChains(driver)\
        .move_to_element(econtinuoUm)\
        .perform()
        time.sleep(2)


