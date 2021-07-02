# 1 - baixa o arquivo pelo navegador, usando biblioteca pyautogui, time e pyperclip

import pyautogui
import time
import pyperclip

pyautogui.PAUSE = 5
pyautogui.press("winleft")
pyautogui.write("chrome")
pyautogui.press("enter")
time.sleep(10)
pyautogui.hotkey('ctrl', 't')
link = 'https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga'
pyperclip.copy(link)
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('enter')
time.sleep(10)
pyautogui.click(x=373, y=272, clicks=2)
pyautogui.click(x=382, y=350)
pyautogui.click(x=1162, y=158)
pyautogui.click(x=934, y=532)
time.sleep(10)

# 2 - abrir a base de dados usando a bibliotecas pandas e soma faturamento e quantidade

import pandas as pd
pyautogui.PAUSE = 5

df = pd.read_excel(r'/home/bravura/Downloads/Vendas - Dez.xlsx')
print(df)

faturamento = df['Valor Final'].sum()
qtde_produtos = df['Quantidade'].sum()
print(f'soma do faturamento: {faturamento}')
print(f'soma da quantidade de produtos: {qtde_produtos}')

# 3 - abrir uma nova aba para acessar o email e enviar um relátorio de vendas
pyautogui.hotkey('ctrl', 't')
pyautogui.write('mail.google.com')
pyautogui.press('enter')
time.sleep(15)

pyautogui.click(x=71, y=184)
time.sleep(10)
email = 'pythonimpressionador+diretoria@gmail.com'
pyperclip.copy(email)
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('tab')
assunto = 'Relatório de vendas de ontem'
pyperclip.copy(assunto)
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('tab')
texto = f"""
Prezados, bom dia

O faturamento de ontem foi de: R$ {faturamento:,.2f}
A quantidade de Produtos foi de: {qtde_produtos:,}

Abs
Carlos Marinho
"""
pyperclip.copy(texto)
pyautogui.hotkey('ctrl', 'v')
pyautogui.hotkey('ctrl', 'enter')
