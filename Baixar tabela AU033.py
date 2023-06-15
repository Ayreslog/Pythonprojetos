#Acessa e baixa tabelas de Perdas AU033 - (leva entre 12 a 15 minutos pra atualizar)
import pyautogui
import pyperclip
import time
import datetime


data= ("17/08/2022")
#data2=datetime.datetime.strptime(data, "%d/%m/%Y")
#print(data2)



#print("{}/{}/{}".format(data2.day,data2.month,data2.year))

#display(data)

pyautogui.PAUSE = 1


#Abaixo irá acessar a tebela de quebras Au033
pyautogui.click(x=91, y=175)
pyperclip.copy("Au033")
pyautogui.hotkey("ctrl","v")
time.sleep(5)

pyautogui.click(x=151, y=310,clicks=2)
time.sleep(5)
#Movimentar o template
pyautogui.click(x=685, y=488)
time.sleep(5)
pyautogui.click(x=617, y=545)
pyautogui.click(x=780, y=484)
#Preenche a data inicial
pyautogui.click(x=646, y=346)
pyautogui.write("01/08/2022")
time.sleep(2)
#Preenche o dia anterior
pyautogui.click(x=645, y=371)
pyautogui.write(data)
time.sleep(2)

#Seleciona o tipo do produto

pyautogui.click(x=867, y=393)
pyautogui.click(x=714, y=470)
time.sleep(2)
pyautogui.click(x=632, y=432)
time.sleep(2)
#Seleciona todas as lojaslog@2022
pyautogui.click(x=416, y=293)
pyautogui.click(x=843, y=504)
time.sleep(10)
#seleciona o tipo de relatório
pyautogui.click(x=706, y=278,clicks=2)
time.sleep(3)
#localiza a posta que irá salvar
pyautogui.click(x=398, y=377)
time.sleep(10)#teste aqui o tempo certo para não travar a posta.
pyautogui.click(x=779, y=389,clicks=2)
pyautogui.click(x=593, y=503)
pyautogui.write("Users")
pyautogui.press("enter")
pyautogui.click(x=593, y=503)
pyautogui.write("ayres.filho")
pyautogui.press("enter")
pyautogui.click(x=593, y=503)
pyautogui.write("Desktop")
pyautogui.press("enter")
pyautogui.click(x=593, y=503)

#preenche o nome do relatório.
pyautogui.write("Quebras")
pyautogui.click(x=964, y=501)
time.sleep(600)#corrigir este cara para 900 home office

#Cica em ok e Fecha o template
pyautogui.click(x=673, y=405)
pyautogui.click(x=851, y=230)
time.sleep(5)
#apaga o relatório
pyautogui.click(x=152, y=173)
pyautogui.press("backspace")
pyautogui.press("backspace")
pyautogui.press("backspace")
pyautogui.press("backspace")
pyautogui.press("backspace")
pyautogui.press("delete")
pyautogui.press("delete")
pyautogui.press("delete")
pyautogui.press("delete")
pyautogui.press("delete")