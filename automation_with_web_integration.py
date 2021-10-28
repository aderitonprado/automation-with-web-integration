#!/usr/bin/env python
# coding: utf-8

# In[46]:


# instalar o selenium
# !pip install selenium

# baixar o webdriver
# Depende do navegador
# chrome: chromedriver
# firefox: geckodriver

# Obs: pesquise pelo webdriver correspondente ao seu navegador

# importar bibliotecas e criar navegador
from selenium import webdriver
from selenium.webdriver.common.keys import Keys

navegador = webdriver.Chrome("chromedriver.exe") # criamos um navegador controlado remotamente

# passo 1: Pegar a cotação do dolar
# - Entrar no site do google
navegador.get("https://www.google.com/")

# - Pesquisar por "cotação dólar"
# para dizer ao navegador aonde está o campo de busca do google, clique com o botão da direita do mouse no campo de busca
# escolha "inspecionar", na guia do inspetor de código, clique no seletor com o botão da direita do mouse
# escolha copiar -> copiar xpath

# cole o xpath como no trecho abaixo
google = navegador.find_element_by_xpath('/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input')
google.send_keys("Cotação dólar")
google.send_keys(Keys.ENTER)

# - Pegar a cotação do dolar e armazenar
cotacao_dolar = navegador.find_element_by_xpath('//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute("data-value")
print("Dólar: ", cotacao_dolar)

# passo 2: Pegar a cotação do euro
# - Entrar no site do google
navegador.get("https://www.google.com/")

# - Pesquisar por "cotação euro"
google = navegador.find_element_by_xpath('/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input')
google.send_keys("Cotação euro")
google.send_keys(Keys.ENTER)

# - Pegar a cotação do euro e armazenar
# copiaremos o valor do atributo data-value, pois possui a informação que queremos
cotacao_euro = navegador.find_element_by_xpath('//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute("data-value")

# exibindo o valor obtido da cotação atual
print("Euro: ", cotacao_euro)

# passo 3: Pegar a cotação do ouro
# - Entrar no site do google
navegador.get("https://www.melhorcambio.com/ouro-hoje")

# - Pegar a cotação do ouro e armazenar
# copiaremos o atributo value, pois possui a informação que queremos
cotacao_ouro = navegador.find_element_by_xpath('//*[@id="comercial"]').get_attribute("value")

# trocaremos o divisor de milhar
cotacao_ouro = cotacao_ouro.replace(",",".")

# exibindo o valor obtido da cotação atual
print("Ouro: ", cotacao_ouro)


# In[47]:


# instalar a lib do pandas com o comando abaixo
# pip install pandas

# passo 4: Importar a base de dados
import pandas as pd
tabela = pd.read_excel(r"D://local-do-seu-arquivo/Produtos.xlsx")

# - visualizar a base de dados (display usado no jupyter, em sua IDE use o print)
display(tabela)

# passo 5: Atualizar a cotação, o preço de compra e o preço de venda

# - atualizar a cotação usando a lógica:
# Localize na tabela a coluna moeda, se sua linha o valor = dolar, atualize a linha da coluna cotação 
# com o valor da cotação do dolar
tabela.loc[tabela["Moeda"] == "Dólar", "Cotação"] = float(cotacao_dolar)
tabela.loc[tabela["Moeda"] == "Euro", "Cotação"] = float(cotacao_euro)
tabela.loc[tabela["Moeda"] == "Ouro", "Cotação"] = float(cotacao_ouro)

# - atualizar o preço de compra = Preço original * cotação
tabela["Preço Base Reais"] = tabela["Preço Base Original"] * tabela["Cotação"]

# - atualizar o preço de venda = Preço de compra * margem
tabela["Preço Final"] = tabela["Preço Base Reais"] * tabela["Margem"]

# - Formatar valores em Real (opcional, já que pode formatar no excel)
# tabela["Preço Final"] = tabela["Preço Final"].map("R$ {:.2f}".format)

display(tabela) # (display usado no jupyter, em sua IDE use o print)


# In[48]:


# passo 6: Exportar o relatorio atualizado
tabela.to_excel(r"D://local-do-seu-arquivo/Produtos Novo2.xlsx", index=False)

# fechar o navegador
navegador.quit()


# In[ ]:




