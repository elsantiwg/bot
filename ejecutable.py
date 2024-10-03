import pandas as pd
import requests
from bs4 import BeautifulSoup
import pywhatkit as pwk
import time
from datetime import datetime
import pytz
import logging

# Configuración del logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Definir la URL de la página donde se realiza la consulta
URL = "http://190.13.96.92/minisitio/Consulta_Numero_Celular.aspx"

# Cargar el archivo Excel con los números
FILE_PATH = 'numeros_portabilidad.xlsx'
df = pd.read_excel(FILE_PATH)

# Crear una sesión para mantener las cookies
session = requests.Session()

# Hacer una solicitud inicial para obtener los valores de VIEWSTATE y EVENTVALIDATION
response = session.get(URL)
soup = BeautifulSoup(response.content, 'html.parser')

# Obtener los valores de VIEWSTATE, VIEWSTATEGENERATOR y EVENTVALIDATION necesarios
viewstate = soup.find("input", {"id": "__VIEWSTATE"}).get("value")
viewstate_generator = soup.find("input", {"id": "__VIEWSTATEGENERATOR"}).get("value")
event_validation = soup.find("input", {"id": "__EVENTVALIDATION"}).get("value")

# Parámetros configurables
MENSAJES_POR_LOTE = 5          # Número de mensajes por lote
TIEMPO_ESPERA_LOTE = 600       # Tiempo de espera entre lotes en segundos (10 minutos)
TIEMPO_ESPERA_ENVIO = 15       # Tiempo de espera para enviar el mensaje en segundos
PAUSA_ENTRE_MENSAJES = 5       # Pausa entre mensajes en segundos

# Zona horaria de Colombia
zona_horaria_colombia = pytz.timezone("America/Bogota")

# Función para obtener el saludo según la hora del día
def obtener_saludo():
    hora_actual = datetime.now(zona_horaria_colombia).hour
    if 0 <= hora_actual < 12:
        return "Buenos días"
    elif 12 <= hora_actual < 18:
        return "Buenas tardes"
    else:
        return "Buenas noches"

# Función para verificar el estado de portabilidad de un número
def verificar_portabilidad(numero):
    payload = {
        '__VIEWSTATE': viewstate,
        '__VIEWSTATEGENERATOR': viewstate_generator,
        '__EVENTVALIDATION': event_validation,
        'txtIngresar': numero,
        'btnIngrasar.x': '0',
        'btnIngrasar.y': '0',
    }
    # Hacer la solicitud POST
    response = session.post(URL, data=payload)
    soup = BeautifulSoup(response.content, 'html.parser')
    mensaje_estado = soup.find("span", {"id": "lblInfoEstadoPortabilidad"}).text.strip()
    return mensaje_estado

# Función para mostrar el temporizador
def mostrar_temporizador(tiempo_restante):
    while tiempo_restante > 0:
        mins, secs = divmod(tiempo_restante, 60)
        temporizador = f"{mins:02d}:{secs:02d}"
        print(f"Tiempo restante para el siguiente lote: {temporizador}", end="\r")
        time.sleep(1)
        tiempo_restante -= 1
    print("")

# Lista para almacenar los resultados de verificación y envío
resultados = []

# Función para procesar un número
def procesar_numero(row):
    numero = str(row['Numero']).strip()
    logging.info(f"Verificando número: {numero}")

    # Verificar el estado de portabilidad
    estado_portabilidad = verificar_portabilidad(numero)
    logging.info(f"Estado de portabilidad: {estado_portabilidad}")

    if estado_portabilidad == "El número de celular no tiene una solicitud de portabilidad en curso.":
        mensaje = obtener_saludo()
        numero = "+57" + numero if not numero.startswith("+") else numero

        try:
            # Enviar el mensaje usando pywhatkit
            logging.info(f"Enviando mensaje a {numero}: {mensaje}")
            pwk.sendwhatmsg_instantly(numero, mensaje, TIEMPO_ESPERA_ENVIO, True)
            estado_envio = "Enviado"
            time.sleep(PAUSA_ENTRE_MENSAJES)
        except Exception as e:
            logging.error(f"Error al enviar mensaje a {numero}: {e}")
            estado_envio = f"Error: {e}"
    else:
        estado_envio = "No enviado - Portabilidad en curso"

    return {'Numero': numero, 'Estado Portabilidad': estado_portabilidad, 'Estado Envío': estado_envio}

# Procesar los números en lotes
for i in range(0, len(df), MENSAJES_POR_LOTE):
    lote = df.iloc[i:i + MENSAJES_POR_LOTE]

    for _, row in lote.iterrows():
        resultado = procesar_numero(row)
        resultados.append(resultado)

    # Esperar entre lotes
    if i + MENSAJES_POR_LOTE < len(df):
        logging.info(f"Esperando {TIEMPO_ESPERA_LOTE // 60} minutos antes de enviar el siguiente lote...")
        mostrar_temporizador(TIEMPO_ESPERA_LOTE)

# Guardar los resultados en un nuevo archivo Excel
df_resultados = pd.DataFrame(resultados)
df_resultados.to_excel('resultados_finales_portabilidad_y_mensajes_lote.xlsx', index=False)

logging.info("Proceso de verificación y envío de mensajes completado.")
