#!/usr/bin/env python
# coding: utf-8

# Esse projeto consiste em realizar operação como criar um código de automação de análise de dados e elaboração de relatórios
# 

# In[1]:


#importei as bibliotecas:
#automação, por exemplo, interação com teclado
import pyautogui 
#congelar o tempo
import time 
#copiar e colar nos campos de texto
import pyperclip

pyautogui.pause = 1
#Passo 1: Abrir uma nova aba

pyautogui.hotkey('ctrl','t')
#Passo 2: Entrar no link do sistema

link = "https://drive.google.com/drive/my-drive"
pyperclip.copy(link)
pyautogui.hotkey('ctrl','v')
pyautogui.press("enter")
#Passo 3: Navegar até o local onde está a nossa base de dados

#Uso muito time.sleep pois meu notebook é bastante antigo, e com isso tem que dar tempo de 
#estabelecer a comunicação com servidor-desktop
time.sleep(10)
pyautogui.click(333, 574, clicks=2)

time.sleep(10)
pyautogui.click(329, 331)

time.sleep(10)
pyautogui.click(1154, 184)

time.sleep(10)
pyautogui.click(968, 627)
#Passo 4: Baixar a planilha de vendas

import pandas as pd 

tabela = pd.read_excel(r"C:\Users\Doug\Downloads\Vendas - Dez.xlsx")
faturamento = tabela["Valor Final"].sum()
quantidade = tabela["Quantidade"].sum()
#Passo 5: Calcular o faturamento e quntidade de produtos vendidos(faturamento)
pyautogui.hotkey("ctrl", "t")
link="https://mail.google.com/mail/u/0/#inbox"
pyperclip.copy(link)
pyautogui.hotkey("ctrl","v")
pyautogui.press("enter")
time.sleep(10)
#Passo 6: Enviar email para a diretoria com a quantidade e o faturamento que calculamos 
time.sleep(10)
pyautogui.click(94, 202)
time.sleep(10)
pyautogui.write("dsdaamorim@gmail.com") 
pyautogui.press("tab")
time.sleep(10)
pyautogui.press("tab")
time.sleep(10)
assunto = "Relatórios de Vendas"
pyperclip.copy(assunto)
pyautogui.hotkey("ctrl","v")
time.sleep(10)
pyautogui.press("tab")
time.sleep(10)
texto_email = f"""
Prezados, bom dia
O faturamento de ontem foi de: R${faturamento:,.2f}
A quantidade foi de: R${quantidade:,.2f}

Abraços, Douglas
"""
pyperclip.copy(texto_email)
pyautogui.hotkey("ctrl","v")
time.sleep(10)
pyautogui.hotkey("ctrl","enter")


# In[ ]:





# In[ ]:




