import pyautogui
import pyperclip
import time
import pandas as pd

# Vamos começar com o roteiro dos vídeos:

# Aula 01- Automação de processos e tarefas. ]

# Passo 01 - Entras no sistema da empresa (no link do drive)

pyautogui.PAUSE = 4
pyautogui.press("win")
pyautogui.write("opera GX")
pyautogui.press("enter")
pyperclip.copy("https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")

# Passo 02 - Navegar até o local do relatorio (entrar na pasta Exportar)

time.sleep(3)
pyautogui.doubleClick(x=440, y=387)

# Passo 03 - Exportar o relatório (Fazer o download)

time.sleep(2)
pyautogui.click(x=469, y=487)
time.sleep(2)
pyautogui.click(x=1713, y=195)
time.sleep(2)
pyautogui.click(x=1502, y=556)

# Passo 04 - Calcular os indicadores

tabela = pd.read_excel(r"C:\Users\guiga\Downloads\Vendas - Dez.xlsx")
print(tabela)

faturamento = tabela ["Valor Final"].sum()
quantidade = tabela["Quantidade"].sum()
print(faturamento)
print(quantidade)

# Passo 05 - Enviar o email para a diretoria

pyautogui.hotkey("ctrl","t")
pyperclip.copy("https://mail.google.com/mail/u/0/#inbox")
pyautogui.hotkey("ctrl","v")
pyautogui.press("enter")
time.sleep(3)
pyautogui.click(x=116, y=196)
time.sleep(3)
pyperclip.copy("henriquesericovm@outlook.com")
pyautogui.hotkey("ctrl","v")

pyautogui.press("tab")

pyperclip.copy("Você recebeu um email de Gavassa")
pyautogui.hotkey("ctrl","v")

pyautogui.press("tab")

texto = f"""Prezado Henrique,
Você foi furtado por Gavassa.

Utilizando Python e várias linhas de código.

Este foi o faturamento da empresa: R${faturamento:,.2f}
E esta foi a quantidade: {quantidade:,}

Este email foi digitado automaticamente!

Att. Gavassa
"""

pyperclip.copy(texto)
pyautogui.hotkey("ctrl","v")

pyautogui.hotkey("ctrl","enter")