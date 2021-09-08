import pyautogui
import time
import pyperclip
import pandas as pd

#pyautogui.displayMousePosition()
pyautogui.PAUSE = 1

#Passo 1
#Abrir uma nova aba
time.sleep(2)
pyautogui.hotkey('ctrl', 't')
#Entrar no link do sistema
link = "https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga"
pyperclip.copy(link)
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('enter')

#Passo 2
time.sleep(5)
pyautogui.click(389, 270, clicks = 2)
time.sleep(2)

#Passo 3
pyautogui.click(401, 337) #clicar no arquivo
pyautogui.click(1713, 157) #clicar nos 3 pontos
pyautogui.click(1525, 561) #clicar no fazer download 
time.sleep(10)

#Passo 4
tabela = pd.read_excel(r'C:\Users\Pichau\Downloads\Vendas - Dez.xlsx')
faturamento = tabela['Valor Final'].sum()
quantidade = tabela['Quantidade'].sum()

#Passo 5
time.sleep(2)
pyautogui.hotkey('ctrl', 't')
#Entrar no link do sistema
link = "https://mail.google.com/mail/u/0/#inbox"
pyperclip.copy(link)
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('enter')
time.sleep(7)
pyautogui.click(33, 170)
pyautogui.write('gustavo.ibis.gb+diretoria@gmail.com')
pyautogui.press('tab')
pyautogui.press('tab')
assunto = 'Relatório de Vendas'
pyperclip.copy(assunto)
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('tab')
texto_email = f"""
Prezados, bom dia

O faturamento de ontem foi de: R${faturamento:,.2f}
A quantidade de produtos foi de: R${quantidade:,.2f}

Abs
"""
pyperclip.copy(texto_email)
pyautogui.hotkey('ctrl', 'v')
pyautogui.hotkey('ctrl', 'enter')