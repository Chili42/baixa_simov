#!/usr/bin/env python
# coding: utf-8

# In[31]:


from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import os
from time import sleep
from bs4 import BeautifulSoup
import pandas as pd
import getpass
from datetime import date
from selenium.webdriver.common.alert import Alert 


# In[32]:


# main
matricula = input('Digite sua matrícula: ')
senha = getpass.getpass('Digite sua senha: ')


# In[33]:


data = date.today()
hoje = data.today().strftime('%d/%m/%Y')
arquivo_base = pd.read_excel(r'derrubar.xlsx')
nubem = arquivo_base['NU_BEM']


# In[34]:


browser = webdriver.Chrome(executable_path='drivers/chromedriver.exe')

url = 'endereço do simov'

browser.get(url)
browser.maximize_window()
form_usuario = browser.find_element_by_id('username')
form_usuario.send_keys(matricula)

form_senha = browser.find_element_by_id('password')
form_senha.send_keys(senha)
passwdbtn = browser.find_element_by_id('kc-login').click()


# In[35]:


try:
    sleep(3)
    browser.switch_to.frame(1)
    browser.switch_to.frame('frmMain')
    browser.find_element_by_xpath('/html/body/form/table/tbody/tr[3]/td[2]/select/option[12]').click()
    browser.find_element_by_name('Image11').click()
    sleep(1)
    browser.get('https://simov.caixa/sistema/asp/alienar/contratacao/cad/venda_online/movacontratacao_cad_cadastro_filtro.asp?QS_sModulo=Incluir&Id=2')
except:
    sleep(5)
    browser.switch_to.frame(1)
    browser.switch_to.frame('frmMain')
    browser.find_element_by_xpath('/html/body/form/table/tbody/tr[3]/td[2]/select/option[12]').click()
    browser.find_element_by_name('Image11').click()
    sleep(1)
    browser.get('https://simov.caixa/sistema/asp/alienar/contratacao/cad/venda_online/movacontratacao_cad_cadastro_filtro.asp?QS_sModulo=Incluir&Id=2')


# In[6]:


for bem in nubem:
    try:
        browser.find_element_by_id('txtNuBem').send_keys(bem)
        browser.find_element_by_name('Image15').click()
        sleep(1)
        browser.find_element_by_xpath("//select[@name='status_contrato']/option[text()='Distrato']").click()
        browser.find_element_by_xpath("//select[@name='NU_MOTIVO_DISTRATO_CONTRATO']/option[text()='CANCELADO-INTERESSE CAIXA']").click()
        browser.find_element_by_xpath("//select[@name='fase_contrato']/option[text()='EM ANÁLISE DE CONFORMIDADE']").click()
        browser.find_element_by_name('data').send_keys(hoje)
        browser.find_element_by_name('Image15').click()
        browser.find_element_by_name('Image19').click()
        Alert(browser).dismiss()
        Alert(browser).accept()
    except:
        browser.get('https://simov.caixa/index.asp')
        sleep(3)
        browser.switch_to.frame(1)
        browser.switch_to.frame('frmMain')
        browser.find_element_by_xpath('/html/body/form/table/tbody/tr[3]/td[2]/select/option[12]').click()
        browser.find_element_by_name('Image11').click()
        browser.get('https://simov.caixa/sistema/asp/alienar/contratacao/cad/venda_online/movacontratacao_cad_cadastro_filtro.asp?QS_sModulo=Incluir&Id=2')
        browser.find_element_by_id('txtNuBem').send_keys(bem)
        browser.find_element_by_name('Image15').click()
        sleep(1)
        browser.find_element_by_xpath("//select[@name='status_contrato']/option[text()='Distrato']").click()
        browser.find_element_by_xpath("//select[@name='NU_MOTIVO_DISTRATO_CONTRATO']/option[text()='CANCELADO-INTERESSE CAIXA']").click()
        browser.find_element_by_xpath("//select[@name='fase_contrato']/option[text()='EM ANÁLISE DE CONFORMIDADE']").click()
        browser.find_element_by_name('data').send_keys(hoje)
        browser.find_element_by_name('Image15').click()
        browser.find_element_by_name('Image19').click()


# In[30]:


browser.get('endereço do simov 2')
sleep(3)
browser.switch_to.frame(1)
browser.switch_to.frame('frmMain')
browser.find_element_by_xpath('/html/body/form/table/tbody/tr[3]/td[2]/select/option[12]').click()
browser.find_element_by_name('Image11').click()
browser.get('https://simov.caixa/sistema/asp/alienar/contratacao/cad/venda_online/movacontratacao_cad_cadastro_filtro.asp?QS_sModulo=Incluir&Id=2')
browser.find_element_by_id('txtNuBem').send_keys(bem)
browser.find_element_by_name('Image15').click()
sleep(1)
browser.find_element_by_xpath("//select[@name='status_contrato']/option[text()='Distrato']").click()
browser.find_element_by_xpath("//select[@name='NU_MOTIVO_DISTRATO_CONTRATO']/option[text()='CANCELADO-INTERESSE CAIXA']").click()
browser.find_element_by_xpath("//select[@name='fase_contrato']/option[text()='EM ANÁLISE DE CONFORMIDADE']").click()
browser.find_element_by_name('data').send_keys(hoje)
browser.find_element_by_name('Image15').click()
browser.find_element_by_name('Image19').click()
browser.switchTo().alert()
alert.getText()


# In[ ]:




