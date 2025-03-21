import cv2
import time
from ultralytics import YOLO

# Carregar o modelo YOLOv8 pre-treinado
model = YOLO('yolov8n.pt')  # 'yolov8n.pt' é a versão leve do modelo

# Iniciar a captura de video da camera
cap = cv2.VideoCapture(0)  # 0 para a camera padrao (frontal)

# Variavel para controle do tempo
tempo_inicial = None
tempo_presente = 0

# Definir a janela em modo tela cheia
cv2.namedWindow('Object Detection', cv2.WND_PROP_FULLSCREEN)
cv2.setWindowProperty('Object Detection', cv2.WND_PROP_FULLSCREEN, cv2.WINDOW_FULLSCREEN)

while True:
    ret, frame = cap.read()
    if not ret:
        break

    # Realizar a deteccao no frame capturado
    results = model(frame)

    # Renderizar as deteccoes no frame
    annotated_frame = results[0].plot()

    # Verificar se uma pessoa foi detectada
    detected_person = False
    for result in results[0].boxes:
        if result.cls == 0:  # Classe 0 é para pessoas no YOLOv8
            detected_person = True
            break

    # Contar o tempo que a pessoa fica em frente a camera
    if detected_person:
        if tempo_inicial is None:
            tempo_inicial = time.time()  # Inicia o contador
        tempo_presente = time.time() - tempo_inicial  # Atualiza o tempo
    else:
        tempo_inicial = None  # Resetar o contador se nao houver pessoa

    # Adicionar o texto com o tempo na imagem
    text_tempo = f"Tempo em frente a camera: {int(tempo_presente)} segundos"
    cv2.putText(annotated_frame, text_tempo, (10, 30), cv2.FONT_HERSHEY_SIMPLEX, 0.7, (0, 255, 255), 2, cv2.LINE_AA)  # Texto amarelo

    # Mostrar o frame com as detecções e tempo
    cv2.imshow('Object Detection', annotated_frame)

    # Encerrar o loop se a tecla 'q' for pressionada
    if cv2.waitKey(1) & 0xFF == ord('q'):
        break

# Liberar a captura e fechar todas as janelas
cap.release()
cv2.destroyAllWindows()
