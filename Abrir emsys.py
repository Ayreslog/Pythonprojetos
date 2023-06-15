#Invocar o EMSYS e logar a senha e usuário
import pyautogui
import pyperclip
import time

pyautogui.PAUSE = 1

#Invoca tecla windows
pyautogui.press("win")
#procura e invoca o EMSYS
pyautogui.write("Go-global")
time.sleep(10)
pyautogui.press("enter")
time.sleep(10)
pyautogui.click(x=852, y=353)
time.sleep(25)
#entra no EMSYS3
pyautogui.click(x=496, y=235,clicks=2)
time.sleep(18)
#Seleciona o banco de dados
pyautogui.click(x=651, y=365)
pyautogui.write("RedeSim")
pyautogui.press("enter")
time.sleep(5)
#Apaga qualquer usuario logado e entra no sistema
pyautogui.click(x=759, y=432)
pyautogui.press("backspace")
pyautogui.press("backspace")
pyautogui.press("backspace")
pyautogui.press("backspace")
pyautogui.press("backspace")
pyautogui.press("backspace")
pyautogui.press("backspace")
pyautogui.press("backspace")
pyautogui.press("backspace")
pyautogui.press("backspace")
pyautogui.press("backspace")
pyautogui.press("backspace")
pyautogui.press("backspace")
pyautogui.press("backspace")
pyautogui.press("backspace")
pyautogui.click(x=759, y=432)
pyautogui.write("ayrisfilho")
pyautogui.click(x=691, y=496)
pyautogui.write("log@2022")
pyautogui.click(x=903, y=519)
time.sleep(8)
pyautogui.click(x=893, y=484)
time.sleep(12)
#expande os menus
pyautogui.click(x=13, y=197)
time.sleep(2)

