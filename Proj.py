#!/usr/bin/env python
# coding: utf-8

# In[ ]:


#Import Bibliotecas
import pandas as pd
import requests
import os
from webdriver_manager.chrome import ChromeDriverManager
from selenium import webdriver
from selenium.webdriver.support.ui import Select
from time import sleep 
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
import pdfkit
import Zip_function


# In[ ]:


#Criando DF a partir do XLS enviado, pulando duas linhas em branco para criar cabeçalho correto
path = '/Users/claudiojunior/Desktop/Ponchi_Technology/Script_Python/VAAC/Emissao_NF/Relatorio_Omie_NFs_Opea.xlsx'
df = pd.DataFrame(pd.read_excel(path,skiprows = [0,1]))


# In[ ]:


#Criando DF a partir da lista de NFS
df_1 = pd.DataFrame(df,columns = ['Código de Verificação','Documento','Contrato','Cliente'])
df_1.rename(columns={'Código de Verificação': 'COD'},inplace = True)

df_1


# In[ ]:


for index,row in df_1.iterrows():
    
    nome_arquivo = ("/"+"NF_"+str(row["Documento"])+"_OPEA_"+ ".pdf")
    Caminho_novo = "/Users/claudiojunior/Desktop/Ponchi_Technology/Script_Python/VAAC/Emissao_NF/NFs_OPEA"
    arquivo = Caminho_novo + nome_arquivo  
    
    
    print('Iniciando Conexão')
    options = webdriver.ChromeOptions() 
    options.add_argument('window-size=400,800')
    options.add_argument('--headless')
    #options.add_argument("user-data-dir=C:/Users/claudio.ponchi/AppData/Local/Google/Chrome/User Data/Default/Accounts")
    
    Chromedriver = '/Users/claudiojunior/Desktop/Ponchi_Technology/Downloads/chromedriver'
    navegador = webdriver.Chrome(executable_path= Chromedriver, options = options)   
    navegador.get('https://nfe.prefeitura.sp.gov.br/publico/verificacao.aspx')
    #pwindow = navegador.current_window_handle
    print('Conexão Estabelecida')
 
   
    
    
    print('Emitindo PDF')
    navegador.find_element("xpath",'//*[@id="ctl00_body_tbCPFCNPJ"]').send_keys('23092592000114')
    sleep(1)
    
    
    navegador.find_element("xpath",'//*[@id="ctl00_body_tbNota"]').send_keys(row["Documento"])
    sleep(1)
    
    
    navegador.find_element("xpath", '//*[@id="ctl00_body_tbVerificacao"]').send_keys(row["COD"])
    sleep(1)
    
    
    navegador.find_element("xpath",'//*[@id="ctl00_body_btVerificar"]').click()
    sleep(1)
    
    
       
    
    get_url = navegador.current_url 
   
    #path_wkthmltopdf = b'C:/Program Files/wkhtmltopdf/bin/wkhtmltopdf.exe'     
    
    
    PDFKIT_CONFIGURATION = pdfkit.configuration()
    pdfkit.from_url(get_url,arquivo,configuration=PDFKIT_CONFIGURATION)
     
     
    navegador.close()  
    print('NF emitida com Sucesso')


# In[ ]:


Zip_function.zip_function(Caminho_novo,'NF_OPEA.zip')


# In[ ]:




