from tkinter import *
from tkinter import messagebox as mb
from tkinter import filedialog
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import navegador
import usuario
import app
import time
import openpyxl
import threading



janela = Tk()
janela.title("MailWave")
janela.geometry("720x600+610+153")
janela.iconbitmap(default="C:\\Users\\elton.santos\\Desktop\\projetos\\emailEXE\\icones\\logo.ico")

img_fundo = PhotoImage(file="C:\\Users\\elton.santos\\Desktop\\projetos\\emailEXE\\imagens\\layout.png")
lab_fundo = Label(janela, image=img_fundo)
lab_fundo.pack()

login_temp_text = ""
senha_temp_text = ""
planilha_temp_text = ""

login = Entry(janela,
              fg="white",
              background="#1E2124",
              bd=2,
              font=("Arimo", 11),
              justify=CENTER,
              )
login.insert(0, login_temp_text)
login.place(x=60, y=213, width=282, height=45)
login.bind("<FocusIn>", lambda event: usuario.set_temp_text(event, usuario, login_temp_text))
login.bind("<FocusOut>", lambda event: usuario.clear_temp_text(event, senha, login_temp_text))

senha = Entry(janela,
              fg="white",
              background="#1E2124",
              bd=2,
              font=("Arimo", 11),
              justify=CENTER,
              show="*",
              )
senha.insert(0, senha_temp_text)
senha.place(x=60, y=293, width=282, height=45)

panilha = Entry(janela,
                fg="#505050",
                background="#1E2124",
                bd=2,
                font=("Arimo", 11),
                justify=CENTER,
                )
panilha.insert(0, planilha_temp_text)
panilha.place(x=60, y=373, width=282, height=45)
panilha.bind("<FocusIn>", lambda event: set_temp_text(event, panilha, planilha_temp_text))
panilha.bind("<FocusOut>", lambda event: clear_temp_text(event, panilha, planilha_temp_text))

butao = Button(janela,
               bd=2,
               fg="white",
               font=("Arimo", 11),
               justify=CENTER,
               text="Entrar",
               background="#8F01EC",
               command=app.teste,
   
               )
butao.place(x=161, y=529, width=80, height=39)

selecionar_planilha = Button(janela, text="Selecionar Planilha", fg="white",  background="#8F01EC")
selecionar_planilha.place(x=125, y=430, width=148, height=40)

janela.mainloop()
