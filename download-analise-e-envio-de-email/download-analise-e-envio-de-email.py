#!/usr/bin/env python
# coding: utf-8

# In[40]:


#Instalação do pyautogui no jupyter
get_ipython().system('pip install pyautogui ')


# In[41]:


import pyautogui as pa
import pyperclip as pc
import pandas

pa.PAUSE = 1
import time

##acessar o navegador e fazer o download da base de dados
pa.hotkey("ctrl", "t")
pc.copy("https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga")
pa.hotkey("ctrl", "v")
pa.press("enter")
time.sleep(2)
pa.click(x=-1911, y=343,clicks=2)
time.sleep(2)
pa.click(x=-1925, y=430)
pa.click(x=-256, y=205)
pa.click(x=-476, y=708)




#Ler o Arquivo e calcular os dados

importar = pandas.read_excel(r"C:\Users\vitor\Downloads\Vendas - Dez.xlsx")
display(importar)
quantidade = importar["Quantidade"].sum()
valor = importar["Valor Final"].sum()

#Acessar o e-mail e enviar o relatório com as informações solicitas (quantidade e valor total de vendas)

pa.hotkey("ctrl", "t")
pc.copy("https://mail.google.com/mail/u/0/#inbox")
pa.hotkey("ctrl", "v")
pa.press("enter")
time.sleep(2)
pa.click(x=-2259, y=210)
time.sleep(2)
pa.write("vap352@gmail.com")
pa.press("tab")
pa.press("tab")
pc.copy("Relatório de vendas")
pa.hotkey("ctrl", "v")
pa.press("tab")
pc.copy(f"""Boa tarde,

Segue o relatório de vendas,

O total vendido foi {quantidade}

Valor arrecadado foi R$ {valor:.2f}

Atenciosamente,

Vitor Alves
""")
pa.hotkey("ctrl","v")
pa.hotkey("ctrl","enter")