#ABRE O EMSYS EM PYWHIN AUTO - PASSO 1 LOGIN
from pywinauto import Application, mouse
import time
import pywinauto.keyboard as keyboard
import pytesseract
import cv2
import numpy as np
import pyautogui

# Configure o caminho do Tesseract, se necessário
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

def check_image_appearance(reference_image_path):
    # Captura a tela inteira
    screenshot = pyautogui.screenshot()
    screenshot = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)

    # Carrega a imagem de referência
    reference_image = cv2.imread(reference_image_path, cv2.IMREAD_UNCHANGED)
    
    # Converte as imagens para escala de cinza
    screenshot_gray = cv2.cvtColor(screenshot, cv2.COLOR_BGR2GRAY)
    reference_image_gray = cv2.cvtColor(reference_image, cv2.COLOR_BGR2GRAY)

    # Realiza a correspondência do modelo
    result = cv2.matchTemplate(screenshot_gray, reference_image_gray, cv2.TM_CCOEFF_NORMED)
    min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(result)

    # Define o limiar de similaridade
    threshold = 0.8  # 80% de correspondência
    if max_val >= threshold:
        print(f"A imagem {reference_image_path} apareceu!")
        return True  # Retorna True se a imagem foi encontrada
    else:
        print(f"A imagem {reference_image_path} não apareceu!")
        return False  # Retorna False se a imagem não foi encontrada

def fallback_process():
    """Executa as etapas quando nenhuma das imagens aparece."""
    print("Executando o processo alternativo...")

    # Coordenadas para o duplo clique e selecionar o EMSYS 3
    x, y = 277, 238  # Coordenadas para clicar no EMSYS 3
    mouse.double_click(coords=(x, y))

    # Aguardar 15 segundos
    time.sleep(25)

    # Coordenadas para digitar o texto
    type_x, type_y = 644, 364  # Coordenadas para onde o texto será digitado
    mouse.move(coords=(type_x, type_y))
    mouse.click(coords=(type_x, type_y))
    keyboard.send_keys("RedeSIM")

    # Coordenadas para o próximo clique
    click_x, click_y = 959, 368
    mouse.click(coords=(click_x, click_y))

    # Coordenadas para digitar o texto 'RPA.LMC'
    type_x1, type_y1 = 643, 429  # Coordenadas para o primeiro texto
    mouse.move(coords=(type_x1, type_y1))
    mouse.click(coords=(type_x1, type_y1))
    keyboard.send_keys("ayrisfilho")

    # Coordenadas para digitar o texto '123'
    type_x2, type_y2 = 645, 495  # Coordenadas para o segundo texto
    mouse.move(coords=(type_x2, type_y2))
    mouse.click(coords=(type_x2, type_y2))
    keyboard.send_keys("123")

    # Coordenadas para o próximo clique
    final_click_x, final_click_y = 897, 525
    mouse.click(coords=(final_click_x, final_click_y))

    final_click_x, final_click_y = 926, 506
    mouse.click(coords=(final_click_x, final_click_y))

    time.sleep(20)

    # Coordenadas para o próximo clique em expandir Menus de Acesso
    final_click_x, final_click_y = 13, 197
    mouse.click(coords=(final_click_x, final_click_y))

def main_process():
    # Iniciar o gg-client.exe com o caminho completo
    app = Application().start(r"C:\Program Files (x86)\GraphOn\GO-Global\Client\gg-client.exe")

    # Aguardar que a primeira janela apareça
    dlg1 = app.window(title_re=".*gg-client.*")

    # Aguardar e interagir com a nova janela que aparece após a inicialização
    dlg2 = app.window(title_re=".*Connection.*")

    # Exemplo: Clicar em um botão na nova janela
    dlg2.ButtonName.click()

    time.sleep(30)

    # Caminho da imagem de referência
    reference_image_path = r'C:\Users\ayres.filho\Imagens\imagem.png'

    # Verificar a primeira imagem
    if check_image_appearance(reference_image_path):
        # Clique em apagar
        pyautogui.click(x=771, y=377)
        # Palavra a ser apagada
        word_to_delete = "@Processos78910"
        num_characters = len(word_to_delete)
        time.sleep(1)

        # Apagar todos os caracteres
        pyautogui.typewrite(['backspace'] * num_characters)
        time.sleep(1)

        # Digita as credenciais do Login
        pyautogui.typewrite('@Processos789', interval=0.1)

        # Clique em Sign in (Entrar)
        pyautogui.click(x=630, y=433)
        time.sleep(30)

        # Caminho da segunda imagem de referência (signin falhou)
        signin_failed_image_path = r'C:\Users\ayres.filho\Imagens\signinfalhou.png'

        # Verificar a segunda imagem
        if check_image_appearance(signin_failed_image_path):
            print("A imagem de erro de login apareceu!")

            # Clicar na posição (696, 435) antes de reiniciar
            pyautogui.click(x=696, y=435)

            # Reiniciar o processo
            print("Reiniciando o prRedeSIMocesso...")
            main_process()  # Reinicia o processo a partir do início
    else:
        # Nenhuma das imagens apareceu, segue para o processo alternativo
        fallback_process()

# Iniciar o processo principal
main_process()
