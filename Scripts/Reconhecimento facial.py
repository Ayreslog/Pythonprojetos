import cv2
import time

# Carregar o classificador Haar Cascade pre-treinado para deteccao de rostos
face_cascade = cv2.CascadeClassifier(cv2.data.haarcascades + 'haarcascade_frontalface_default.xml')

# Iniciar a captura de video da camera
cap = cv2.VideoCapture(0)  # 0 para a camera padrao (frontal)

# Variaveis para controle do tempo
tempo_inicial = None
tempo_presente = 0

# Definir a janela em modo tela cheia
cv2.namedWindow('Face Detection', cv2.WND_PROP_FULLSCREEN)
cv2.setWindowProperty('Face Detection', cv2.WND_PROP_FULLSCREEN, cv2.WINDOW_FULLSCREEN)

while True:
    # Capturar frame por frame
    ret, frame = cap.read()
    
    # Converter o frame para escala de cinza para deteccao de faces
    gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
    
    # Detectar faces no frame
    faces = face_cascade.detectMultiScale(gray, scaleFactor=1.3, minNeighbors=5)
    
    # Se houver deteccao de pelo menos 1 rosto ("tigre")
    if len(faces) > 0:
        if tempo_inicial is None:
            tempo_inicial = time.time()  # Iniciar o contador ao detectar o primeiro rosto
        tempo_presente = time.time() - tempo_inicial  # Atualizar o tempo em frente à camera
    else:
        tempo_inicial = None  # Resetar o contador se nenhum rosto for detectado
    
    # Desenhar um retangulo ao redor de cada rosto detectado
    for (x, y, w, h) in faces:
        cv2.rectangle(frame, (x, y), (x+w, y+h), (255, 0, 0), 2)
    
    # Contar o numero de "tigres" detectados
    num_tigres = len(faces)
    
    # Adicionar o texto na imagem com a contagem de tigres (cor: amarelo)
    text_tigres = f"Numero de pessoas detectadas: {num_tigres}"
    cv2.putText(frame, text_tigres, (10, 30), cv2.FONT_HERSHEY_SIMPLEX, 1, (0, 255, 255), 2, cv2.LINE_AA)  # Amarelo
    
    # Adicionar o texto na imagem com o tempo em frente à camera (cor: cyan)
    text_tempo = f"Tempo em frente a camera: {int(tempo_presente)} segundos"
    cv2.putText(frame, text_tempo, (10, 70), cv2.FONT_HERSHEY_SIMPLEX, 0.7, (255, 255, 0), 2, cv2.LINE_AA)  # Cyan
    
    # Mostrar o frame resultante com a contagem de tigres e tempo
    cv2.imshow('Face Detection', frame)
    
    # Encerrar o loop se a tecla 'q' for pressionada
    if cv2.waitKey(1) & 0xFF == ord('q'):
        break

# Liberar a captura e fechar todas as janelas
cap.release()
cv2.destroyAllWindows()
