import pyautogui
import pyperclip
import time

time.sleep(10)

while True:
 #1° LOOP
 pyautogui.PAUSE = 2

 #duplo click na loja
 pyautogui.click(x=263, y=136,clicks=2)
 #duplo click nos itens
 pyautogui.click(x=291, y=450,clicks=2)
 pyautogui.click(x=793, y=440)


 #clica em AUDITORIA
 pyautogui.click(x=491, y=690)
 #clica em OK
 pyautogui.click(x=807, y=442)

 pyautogui.PAUSE = 2

 #2° LOOP
 #duplo click na loja
 pyautogui.click(x=263, y=136,clicks=2)
 #duplo click nos itens
 pyautogui.click(x=291, y=450,clicks=2)

 #clica em AUDITORIA
 pyautogui.click(x=491, y=690)
 #clica em OK
 pyautogui.click(x=807, y=442)
 pyautogui.click(x=793, y=440)

 #REPROVA
 pyautogui.click(x=398, y=695)
 #CLICA EM OK PÓS REPROVAR
 pyautogui.click(x=771, y=444)

