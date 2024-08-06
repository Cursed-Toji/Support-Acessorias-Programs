from selenium import webdriver
from selenium.webdriver.support.select import Select 
from selenium.webdriver import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import time


class bootAcessorias:
    def __init__(self):
        self.options = webdriver.ChromeOptions()
        self.driver = webdriver.Chrome()



    def login(self):
        driver = self.driver
        driver.implicitly_wait(10)
        driver.maximize_window()
        driver.get("https://app.acessorias.com")