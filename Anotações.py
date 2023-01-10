

# Vamos começar com o roteiro dos vídeos:

# Aula 01- Automação de processos e tarefas.

#-------------------------------------------------------------------------------------------------

# Importações de Bibliotecas:
import pyautogui
import pyperclip
import time

# Haverá momentos que será mais comum utilizar apelidos em importações.
# Para apelidar uma importação, utilize "as Apelido" após a importação:
import pandas as pd

#-------------------------------------------------------------------------------------------------

# Passo 01 - Entrar no sistema da empresa (no link do drive).

# Para determinar uma pausa entre as ações:
pyautogui.PAUSE = 4

# Para determinar o pressionamento de uma tecla:
pyautogui.press("win")

# Para escrever algo com o teclado:
pyautogui.write("opera GX")

# Para determinar o pressionamento de uma tecla:
pyautogui.press("enter")

# O Pyautogui, tem uma restrição com caracteres especiais, portanto é necessário utilizar o Pyperclip:
# Utilizando o Pyperclip para copiar o link do drive que devemos entrar e adquirir as informações necessárias.
pyperclip.copy("https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga")

# Para teclar algum atalho, tal como "Ctrl+V" para colar algo previamente copiado.
pyautogui.hotkey("ctrl", "v")

# Para determinar o pressionamento de uma tecla:
pyautogui.press("enter")

#-------------------------------------------------------------------------------------------------

# Passo 02 - Navegar até o local do relatorio (entrar na pasta Exportar)

# Para determinar uma pausa neste momento:
time.sleep(3)

# Para clicar duas vezes na posição determinada:
pyautogui.doubleClick(x=440, y=387)

# Caso queira apertar com o botão direito:

# pyautogui.rightClick(x=Posicao1, y=Posicao2)

#-------------------------------------------------------------------------------------------------

# Passo 03 - Exportar o relatório (Fazer o download)

time.sleep(2)
pyautogui.click(x=469, y=487)
time.sleep(2)
pyautogui.click(x=1713, y=195)
time.sleep(2)
pyautogui.click(x=1502, y=556)

#-------------------------------------------------------------------------------------------------

# Passo 04 - Calcular os indicadores

# Determine uma variável, para guardar as informações da tabela baixada:
tabela = pd.read_excel(r"C:\Users\guiga\Downloads\Vendas - Dez.xlsx")
# Aqui utiliza-se a bibliteca Pandas, que servirá para ler as informações do arquivo Excel (Ela lê não só Excel, mas como diversos outros tipos de arquivos)
# Ao ditar o caminho do arquivo, é necessário colocar o r antes, para que não tenha erros de comando por conta dos \U, \g, \D e etc.

# Exiba a tabela no terminal, para controle.
print(tabela)

# Calcule os indicadores utilizando variáveis

# Para selecionar uma coluna especifica, utilize ["NomeDaColuna"]
faturamento = tabela ["Valor Final"].sum()

# E para somar todos os itens da coluna utilize .sum()
quantidade = tabela["Quantidade"].sum()

print(faturamento)
print(quantidade)

#-------------------------------------------------------------------------------------------------


# Passo 05 - Enviar o email para a diretoria

# Para teclar algum atalho, tal como "Ctrl+T" para abrir uma nova guia no navegador.
pyautogui.hotkey("ctrl","t")

# O Pyautogui, tem uma restrição com caracteres especiais, portanto é necessário utilizar o Pyperclip:
# Utilizando o Pyperclip para copiar o link do drive que devemos entrar e adquirir as informações necessárias.
pyperclip.copy("https://mail.google.com/mail/u/0/#inbox")

# Para teclar algum atalho, tal como "Ctrl+V" para colar algo previamente copiado.
pyautogui.hotkey("ctrl","v")

# Para determinar o pressionamento de uma tecla:
pyautogui.press("enter")
time.sleep(3)
pyautogui.click(x=116, y=196)
time.sleep(3)
pyperclip.copy("henriquesericovm@outlook.com")
pyautogui.hotkey("ctrl","v")

# Para determinar o pressionamento de uma tecla:
pyautogui.press("tab")

# O Pyautogui, tem uma restrição com caracteres especiais, portanto é necessário utilizar o Pyperclip:
# Utilizando o Pyperclip para copiar o link do drive que devemos entrar e adquirir as informações necessárias.
pyperclip.copy("Você recebeu um email de Gavassa")

# Para teclar algum atalho, tal como "Ctrl+V" para colar algo previamente copiado.
pyautogui.hotkey("ctrl","v")

# Para determinar o pressionamento de uma tecla:
pyautogui.press("tab")

# Determine uma variável para adicionar o texto do email,
# e para escrever o email com todas as quebras de linhas, em uma variável só:
# Utilize Três aspas (duplas ou simples) e, escreva o texto.
texto = f"""Prezado Henrique,
Você foi furtado por Gavassa.

Utilizando Python e várias linhas de código.

Este foi o faturamento da empresa: R${faturamento:,.2f}
E esta foi a quantidade: {quantidade:,}

Este email foi digitado automaticamente!

Att. Gavassa
"""

# Para adicionar uma variável, dentro de outra, porém sendo dinâmica, coloque-a dentro de chaves -> {}
# Porém, também terá que adicionar um f, antes das aspas, para que o texto seja formatado.

# Um exemplo de formatações para a variável dinâmica:
# R${faturamento:,.2f} -> :, servirá para separar os números por vírgulas, a cada três algarismos.

# R${faturamento:,.2f} -> .2f irá separar os dois últimos algarismos com um ponto, tornando o número inteiro em float.

# Copiar a variável do texto do email, para utilizar posteriormente no código
pyperclip.copy(texto)

# Para teclar algum atalho, tal como "Ctrl+V" para colar algo previamente copiado.
pyautogui.hotkey("ctrl","v")

# Para teclar algum atalho, tal como "Ctrl+V" para colar algo previamente copiado.
pyautogui.hotkey("ctrl","enter")