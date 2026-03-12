# %% [markdown]
# # Automação de Sistemas e Processos com Python
# 
# ### Desafio:
# 
# Todos os dias, o nosso sistema atualiza as vendas do dia anterior.
# O seu trabalho diário, como analista, é enviar um e-mail para a diretoria, assim que começar a trabalhar, com o faturamento e a quantidade de produtos vendidos no dia anterior
# 
# E-mail da diretoria: seugmail+diretoria@gmail.com<br>
# Local onde o sistema disponibiliza as vendas do dia anterior: https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga
# 
# Para resolver isso, vamos usar o pyautogui, uma biblioteca de automação de comandos do mouse e do teclado
# 
# 
# Referência do pyautogui: https://pyautogui.readthedocs.io/en/latest/quickstart.html

# %%
# AUTOMAÇÃO DE SISTEMAS E PROCESSOS
# Caulcular os indicadores ( Faturamento e a quantidade dos produtos vendidos)

import openpyxl
import pandas
import pyautogui
import webbrowser
import time

# pyautogui.click -> clicar com o mouse
# pyautogui.write -> escrever um texto
# pyautogui.press -> aperta 1 tecla
# pyautogui.hotkey -> combinação de teclas

# Passo a passo do projeto 
# Passo 1: Entrar no sistema da empresa ( LINK DO DRIVE)
# Passo 2: Navegar no sistema para encontrar a base de dados (o arquivo do drive)
# Passo 3: Exportar a base de dados ( baixar o arquivo)
# Passo 4: Caulcular os indicadores ( Faturamento e a quantidade dos produtos vendidos)
# Passo 5: Enviar o relatprio por e-mail

# Passo a passo do projeto 
# Passo 1: Entrar no sistema da empresa ( LINK DO DRIVE)
# Abrir o chrome


# Abrie o link do sistema
link = "https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga"
webbrowser.open(link)
time.sleep(7)
#pyautogui.write(link)
#pyautogui.press(enter)
time.sleep(7)





# %% [markdown]
# ### Vamos agora ler o arquivo baixado para pegar os indicadores
# 
# - Faturamento
# - Quantidade de Produtos

# %%
# Passo 2: Navegar no sistema para encontrar a base de dados (o arquivo do drive)
# abrir o arquivo
pyautogui.click(x=370, y=364, clicks=2)
time.sleep(6)
# Passo 3: Exportar a base de dados ( baixar o arquivo)
# Clocar com obotão diretio do mouse encima do arquivo
pyautogui.rightClick(x=383, y=379)
# Ir para opçao de baixar o arquivo e clicar
time.sleep(9)
pyautogui.click(x=434, y=402)


# %% [markdown]
# ### Vamos agora enviar um e-mail pelo gmail

# %%
# Passo 4: Caulcular os indicadores ( Faturamento e a quantidade dos produtos vendidos)
# Abrir a base de dados que foi baixada
# Ler a informacões da base de dados
# Ver as infomacões da base de dados
# Somar todos os faturamento de todos os produtos da base de dados (Somar a coluna de valor final)
# Somar todos os produtos da base de dados (Somar quantidades de produtos vendidos) = (Somar a coluna de quantidade)
import pandas
# Abrir o arquivo com o caminho do arquivo 
caminho_arquivo = r"C:\Users\Sheila Lima\Downloads\Vendas - Dez.xlsx"
# Ler o arquivo
tabela_vendas = pandas.read_excel(caminho_arquivo)
# Ver as infomacões
print(tabela_vendas)
# Somar o faturamento = somar a coluna de valor final
faturamento = tabela_vendas["Valor Final"].sum()
# Somar a coluna de quantidade de protudos vendas
qtde_produtos = tabela_vendas["Quantidade"].sum()
print(faturamento)
print(qtde_produtos)



# %%
# Passo 5: Enviar o relatprio por e-
import pyperclip
  # abrir o e-mail nonavegador
pyautogui.hotkey("ctrl", "t")
time.sleep(10)

  #fazer logn no e-mail
pyautogui.write("gmail.com")
pyautogui.press("enter")
time.sleep(13)  
  #escrever o e-mail e enviar
    # ecrever novo e-mail
pyautogui.click(x=77, y=221)
time.sleep(6)
    # digitar o e-mail do destinarario
#pyautogui.click(x=842, y=295)
pyautogui.write("emailnaotem@hotmail.com")
pyautogui.press("tab") # celecionar o email
    # digitar o assunto
pyautogui.press("tab")    
#pyautogui.click(x=869, y=333)
pyperclip.copy("Relatório de vendas")
pyautogui.hotkey("Ctrl", "v")
    # escrever o corpo do e-mail
pyautogui.press("tab")  
texto_email = f""" 
Prezados diretores,

segui o relatorio de vendas de hoje.

Faturamento foi de: R${faturamento:,.2f}
Qantidade de produdos venditos foi: {qtde_produtos:,}

dividas, entrar em contato, estarei a disposição

Ass: Paulo R Queiroz
Codigo Python
                """
pyperclip.copy(texto_email)
pyautogui.hotkey("Ctrl", "v")
#pyautogui.write(texto_email)
   # Enviar o emaol
pyautogui.click(x=847, y=687)   


# %% [markdown]
# #### Use esse código para descobrir qual a posição de um item que queira clicar
# 
# - Lembre-se: a posição na sua tela é diferente da posição na minha tela

# %%
#import time
#import pyautogui
#time.sleep(13)
#print(pyautogui.position())



