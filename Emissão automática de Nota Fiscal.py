#!/usr/bin/env python
# coding: utf-8

# In[2]:


from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By

options = webdriver.ChromeOptions()
options.add_experimental_option("prefs", {
  "download.default_directory": r"C:\Usuários\arthu\downloads\Test",
  "download.prompt_for_download": False,
  "download.directory_upgrade": True,
  "safebrowsing.enabled": True
})


servico = Service(ChromeDriverManager().install())
navegador = webdriver.Chrome(service=servico, chrome_options=options )


# In[3]:


import os #entrar na pagina login (no caso é login.html)

caminho = os.getcwd()
arquivo = caminho + r"\login.html"
navegador.get(arquivo)


# ### Importando a tabela das notas a ser emitidas ###

# In[4]:


import pandas as pd

notas_df = pd.read_excel(r"C:\Users\arthu\exercicio01\NotasEmitir.xlsx")
display(notas_df)


# ### Fazendo Login ###

# In[5]:


#preenchendo login e senha
navegador.find_element(By.XPATH, '/html/body/div/form/input[1]').send_keys('Arthur')
navegador.find_element(By.XPATH, '/html/body/div/form/input[2]').send_keys('1234')
navegador.find_element(By.XPATH, '/html/body/div/form/button').click()


# ### Fazendo um for para enviar todas a notas ###

# In[6]:


from time import sleep

for linha in notas_df.index:
    #preencher dados 
    navegador.find_element(By.XPATH, '//*[@id="nome"]').send_keys(notas_df.loc[linha,"Cliente"])
    #endereço
    navegador.find_element(By.NAME, 'endereco').send_keys(notas_df.loc[linha,'Endereço'])
    #bairro
    navegador.find_element(By.NAME, 'bairro').send_keys(notas_df.loc[linha,'Bairro'])
    #cidade
    navegador.find_element(By.NAME, 'municipio').send_keys(notas_df.loc[linha,'Municipio'])
    #cep
    navegador.find_element(By.NAME, 'cep').send_keys(str(notas_df.loc[linha,'CEP']))
    #estado
    navegador.find_element(By.NAME, 'uf').send_keys(notas_df.loc[linha,'UF'])
    #cpf
    navegador.find_element(By.NAME, 'cnpj').send_keys(str(notas_df.loc[linha,'CPF/CNPJ']))
    #inscrição estadual
    navegador.find_element(By.NAME, 'inscricao').send_keys(str(notas_df.loc[linha,'Inscricao Estadual']))
    #descrição
    navegador.find_element(By.NAME, 'descricao').send_keys(notas_df.loc[linha,'Descrição'])
    #quantidade
    navegador.find_element(By.NAME, 'quantidade').send_keys(str(notas_df.loc[linha,'Quantidade']))
    #valor unitario
    navegador.find_element(By.NAME, 'valor_unitario').send_keys(str(notas_df.loc[linha,'Valor Unitario']))
    #valor total
    navegador.find_element(By.NAME, 'total').send_keys(str(notas_df.loc[linha,'Valor Total']))
    #clicar em emitir nota fiscal
    navegador.find_element(By.CLASS_NAME, 'registerbtn').click()
    
    sleep(0.2)
    
    #recarregando a página
    navegador.refresh()
    

    
    
  




# In[ ]:


navegador.quit()

