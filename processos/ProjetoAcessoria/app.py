from selenium import webdriver
from selenium.webdriver.support.select import Select 
from selenium.webdriver import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import time
import navegador
import usuario


def teste():
    # Inicializa objetos, passando o atributo "driver" do objeto boot para o usuário.
    boot = navegador.bootAcessorias()
    user = usuario.usuarioAcessorias(boot.driver)

    # Métodos dos objetos para abrir instância de browser e fazer login
    boot.login()
    user.fazerLogin()
    user.marcarDepartamento()


