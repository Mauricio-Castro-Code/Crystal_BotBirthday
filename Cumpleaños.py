import pandas as pd
import pyautogui as pg
import time
import webbrowser
from datetime import datetime
import re  # Importar para trabajar con expresiones regulares
import os


def normalizar_numero(numero):
    """Normaliza un número de teléfono eliminando caracteres no numéricos."""
    if isinstance(numero, str):
        # Usar expresión regular para eliminar todo lo que no sea un dígito
        return re.sub(r'\D', '', numero)
    elif isinstance(numero, (int, float)):
        # Convertir números a string y asegurarse de que estén formateados como enteros
        return str(int(numero))
    else:
        return None  # Devuelve None si no es un formato válido


def enviar_mensajes_whatsapp(numero, mensaje, imagen_ruta):
    """Enviar mensaje y adjuntar imagen por WhatsApp Web."""
    url = f"https://web.whatsapp.com/send?phone={numero}&text={mensaje}"
    webbrowser.open(url)
    time.sleep(13)  # Tiempo para que WhatsApp Web cargue
    # Adjuntar imagen
    pg.click(712, 955)  # Abre el menú de adjuntar (puedes ajustar las coordenadas)
    time.sleep(2)
    pg.click(825, 520)  # Hace clic en "Fotos y videos" (ajusta según sea necesario)
    time.sleep(8)
    # Escribir la ruta de la imagen
    pg.write(imagen_ruta)
    time.sleep(2)
    pg.press('enter')
    time.sleep(2)
    # Enviar el mensaje
    pg.press('enter')
    # Cerrar la pestaña
    time.sleep(5)
    pg.hotkey('ctrl', 'w')  # 'ctrl' + 'w' en Windows
    pg.press('enter')


# Ruta de la imagen a enviar esto es SOLO UN EJEMPLO
imagen_ruta = "checo1.jpg"

# Leer el archivo Excel EJEMPLO, usando columnas B, C y D (Número, Nombre, FechaNacimiento) y saltando las primeras dos filas
#df = pd.read_excel('datos.xlsx', usecols='B:D', skiprows=2)
# Obtener la ruta del directorio donde se encuentra el script ejecutable
directorio = os.path.dirname(__file__)

# Crear la ruta completa al archivo Excel EJEMPLO
ruta_datos = os.path.join(directorio, 'datos.xlsx')

# Leer el archivo Excel usando la ruta relativa
df = pd.read_excel(ruta_datos, usecols='B:D', skiprows=2)

# Renombrar las columnas para facilitar el acceso
df.columns = ['Numero', 'Nombre', 'FechaNacimiento']

# Normalizar los números de teléfono
df['Numero'] = df['Numero'].apply(normalizar_numero)

# Asegurarse de que las fechas sean correctamente interpretadas
df['FechaNacimiento'] = pd.to_datetime(df['FechaNacimiento'], errors='coerce', dayfirst=True)

# Verificar que las fechas se leyeron correctamente
#print(df.head())

# Obtener fecha actual (día y mes)
fecha_actual = datetime.now().strftime('%d/%m')

# Conjunto para registrar números a los que ya se les ha enviado un mensaje
numeros_mensajes_enviados = set()

# Filtrar clientes con cumpleaños hoy
for _, row in df.iterrows():
    nombre = row.get('Nombre')
    numero = row.get('Numero')
    fecha_nacimiento = row.get('FechaNacimiento')
    # Validar datos
    if pd.isna(nombre) or pd.isna(numero) or pd.isna(fecha_nacimiento):
        print(f"Datos faltantes para la fila: {row}")
        continue
    # Verificar si el cumpleaños coincide
    fecha_nacimiento_str = fecha_nacimiento.strftime('%d/%m')
    if fecha_nacimiento_str == fecha_actual:
        # Verificar si el número ya ha recibido un mensaje
        if numero in numeros_mensajes_enviados:
        #print(f"El número {numero} ya recibió un mensaje. Saltando...")
            continue
        # Crear mensaje y enviar
        mensaje = f"""¡Felicidades {nombre} en tu cumpleaños!

Nos encanta ser parte de tus momentos especiales, y en esta ocasión tan importante, queremos celebrarlo contigo. 
Por eso, durante todo el mes de tu cumpleaños, te ofrecemos un 10% de descuento en todos tus pedidos.

Gracias por confiar en nosotros para tus eventos, ¡esperamos seguir creando recuerdos inolvidables juntos!

Con cariño,
*El equipo de ALQUILADORA CRYSTAL*"""
        enviar_mensajes_whatsapp(numero, mensaje, imagen_ruta)
        # Registrar el número como ya notificado
        numeros_mensajes_enviados.add(numero)
        print(f"{nombre}  ha sido felicitado por el numero: ({numero}).")




