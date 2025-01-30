import pyautogui
import time
#Codigo PARA CHECAR COORDENADAS DE LA PANTALLA
try:
    while True:
        x, y = pyautogui.position()  # Obtiene las coordenadas del cursor
        print(f"Coordenadas del cursor: X={x}, Y={y}")
        time.sleep(1)  # Pausa de 1 segundo para actualizar la posición
except KeyboardInterrupt:
    print("\nPrograma terminado.")