from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import os

script_directory = os.path.dirname(os.path.abspath(__file__))
os.chdir(script_directory)

driver_path = r"C:\Users\luzso\Downloads\chromedriver_win32\chromedriver.exe"
brave_path = r"C:\Users\luzso\AppData\Local\BraveSoftware\Brave-Browser\Application\brave.exe"

option = webdriver.ChromeOptions()

servico = Service(ChromeDriverManager().install())
option.binary_location = brave_path

option.add_argument("--enable-chrome-browser-cloud-management")
option.add_argument(r"user-data-dir=C:\Users\luzso\AppData\Local\BraveSoftware\Brave-Browser\User Data\Profile Selenium")


option.add_experimental_option("prefs", {
    "download.default_directory" : script_directory,
    "download.prompt_for_download" : False
})

option.add_experimental_option("detach", True)
navegador = webdriver.Chrome(service=servico, options=option)

#************************************************************************************************
import urllib
import time

navegador.get('https://web.whatsapp.com/')
WebDriverWait(navegador, 20).until(EC.presence_of_element_located((By.ID, 'side'))) #o elemento que me diz se a tela ja carregou
time.sleep(1) # só uma garantia

tabela = pd.read_excel("Envios.xlsx") 

for linha in tabela.index :
    #enviar uma mensagem para a pessoa
    nome = tabela.loc[linha, 'nome']
    mensagem = tabela.loc[linha, "mensagem"]
    arquivo = tabela.loc[linha, "arquivo"]
    telefone = tabela.loc[linha, "telefone"]


    texto = urllib.parse.quote(mensagem)

    #enviar a mensagem
    if telefone != "N" :
        link = f"https://web.whatsapp.com/send?phone={telefone}&text={texto}"
        navegador.get(link)
        #esperar o zap zap carregar --> esperar um elemento que so existe na tela aparecer
        WebDriverWait(navegador, 20).until(EC.presence_of_element_located((By.ID, 'side'))) #o elemento que me diz se a tela ja carregou
        time.sleep(3) #garantia 
    else :
        navegador.find_element(By.XPATH, '//*[@id="side"]/div[1]/div/div[2]/div[2]/div/div[1]').send_keys(nome)  #caixa de texto
        if navegador.find_element(By.XPATH, '//*[@id="pane-side"]/div[1]/div/div').get_attribute('aria-rowcount') == "1" : #se so tiver um resultado na pesquisa
            time.sleep(1)
            navegador.find_element(By.XPATH, '//*[@id="pane-side"]/div[1]/div/div/div[1]').click() #clicar no grupo
            time.sleep(1)
            navegador.find_element(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div[1]').send_keys(mensagem)




    #verificar se o numero é invalido
    if len(navegador.find_elements(By.XPATH, '//*[@id="app"]/div/span[2]/div/span/div/div/div/div/div/div[1]')) < 1 :
        #enviar a mensagem
        print(len(navegador.find_elements(By.CLASS_NAME, 'message-out focusable-list-item _1AOLJ _1jHIY')))
        navegador.find_element(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[2]/button/span').click()
        #anexar arquivo
        if arquivo != "N" :
            navegador.find_element(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[1]/div[2]/div/div').click() #abre as opcoes
            navegador.find_element(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[1]/div[2]/div/span/div/ul/div/div[1]/li/div/input').send_keys(rf"{script_directory}\arquivos\{arquivo}") #anexa o arquivo
            time.sleep(3)
            navegador.find_element(By.XPATH, '//*[@id="app"]/div/div[2]/div[2]/div[2]/span/div/span/div/div/div[2]/div/div[2]/div[2]/div/div').click() #envia
        time.sleep(3)
