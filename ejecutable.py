import pandas as pd
import requests
from bs4 import BeautifulSoup
import pywhatkit as pwk
import time
from datetime import datetime
import pytz
import logging
import re

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
MENSAJES_POR_LOTE = 5           # Número de mensajes por lote
TIEMPO_ESPERA_LOTE = 600        # Tiempo de espera entre lotes en segundos (10 minutos)
TIEMPO_ESPERA_ENVIO = 15        # Tiempo de espera para enviar el mensaje en segundos
PAUSA_ENTRE_MENSAJES = 5        # Pausa entre mensajes en segundos
MAX_REINTENTOS = 3              # Número máximo de reintentos
PAUSA_INCREMENTAL = 5           # Incremento de la pausa después de cada reintento

# Contadores globales
total_enviados = 0
total_fallidos = 0
total_no_portables = 0

# Zona horaria de Colombia
zona_horaria_colombia = pytz.timezone("America/Bogota")

# Mensajes de campaña
mensaje_promocion = """Dile SÍ al mejor PLAN con 60GB 
¡Este plan tiene los mejores beneficios para ti! 😱: 

Cargo básico de $29.990

😱 Primer Mes Gratis
😱 30% de descuento en la factura mensual del mes 2 al 4  20.990 🫂 Podrás compartir hasta 35GB de tus 60GB.
🤳 Facebook, Messenger, WhatsApp, Instagram, Waze, Twitter, Spotify, Deezer, Nequi y DaviPlata totalmente ilimitadas en tu plan.
📞 Minutos y SMS Ilimitados a todo destino nacional.
📲 500 Minutos a destinos internacionales, (Canadá, EE.UU y Puerto Rico)

Recuerda: Los beneficios no incluyen la licencia de cada aplicación, aplica solo para portabilidad
💙#VamosConToda💙"""

mensaje_confirmacion = """Deseas realizar el proceso de portabilidad? Digita *Si* para continuar, *No* en caso tal de que no deseas realizar el proceso de portabilidad ❤️"""

mensaje_datos = """¡Genial!😊 Para continuar con el proceso es necesario los siguientes datos personales: 
📱Número donde deseas activar el servicio: 
🪪Nombre completo: 
🪪Cédula:
📆Fecha de expedición de la cédula: 
📧Correo: 
🏡Dirección: 
🏘️Barrio y ciudad:"""

# Función para obtener el saludo según la hora del día
def obtener_saludo():
    hora_actual = datetime.now(zona_horaria_colombia).hour
    if 0 <= hora_actual < 12:
        return "Buenos días"
    elif 12 <= hora_actual < 18:
        return "Buenas tardes"
    else:
        return "Buenas noches"

# Función para verificar el estado de portabilidad con reintentos
def verificar_portabilidad(numero):
    intentos = 0
    while intentos < MAX_REINTENTOS:
        try:
            payload = {
                '__VIEWSTATE': viewstate,
                '__VIEWSTATEGENERATOR': viewstate_generator,
                '__EVENTVALIDATION': event_validation,
                'txtIngresar': numero,
                'btnIngrasar.x': '0',
                'btnIngrasar.y': '0',
            }
            response = session.post(URL, data=payload)
            soup = BeautifulSoup(response.content, 'html.parser')
            mensaje_estado = soup.find("span", {"id": "lblInfoEstadoPortabilidad"}).text.strip()
            return mensaje_estado
        except requests.exceptions.Timeout:
            logging.warning(f"Timeout al verificar el número {numero}. Reintentando...")
        except requests.exceptions.ConnectionError:
            logging.error(f"Error de conexión al verificar el número {numero}.")
            return "Error de conexión"
        except Exception as e:
            logging.warning(f"Error en la verificación de portabilidad para {numero}: {e}")
        
        intentos += 1
        time.sleep(PAUSA_INCREMENTAL * intentos)  # Aumentar la pausa con cada intento
    return "Error en la verificación después de varios intentos"

# Función para mostrar el temporizador
def mostrar_temporizador(tiempo_restante):
    for tiempo in range(tiempo_restante, 0, -1):
        mins, secs = divmod(tiempo, 60)
        temporizador = f"{mins:02d}:{secs:02d}"
        print(f"Esperando {temporizador} para el siguiente lote...", end="\r")
        time.sleep(1)
    print("")

# Función para validar números de teléfono (colombianos)
def validar_numero(numero):
    numero = str(numero).strip()
    return len(numero) == 10 and numero.isdigit()

# Función para enviar mensajes masivos
def enviar_mensaje(numero, mensaje):
    try:
        pwk.sendwhatmsg_instantly(numero, mensaje, TIEMPO_ESPERA_ENVIO, True)
        time.sleep(PAUSA_ENTRE_MENSAJES)
        return True
    except Exception as e:
        logging.error(f"Error al enviar mensaje a {numero}: {e}")
        return False

# Función para analizar y responder a los mensajes recibidos durante el tiempo de espera
def analizar_mensajes_entrantes():
    mensajes = pwk.get_chat()  # Obtener los mensajes de los chats abiertos
    for mensaje in mensajes:
        numero = mensaje['number']
        texto = mensaje['message'].lower()

        # Revisar si alguien ha respondido "si" o "no"
        if texto == "si":
            logging.info(f"El contacto {numero} respondió 'Si', enviando mensaje de solicitud de datos.")
            enviar_mensaje(numero, mensaje_datos)
        elif texto == "no":
            logging.info(f"El contacto {numero} respondió 'No', sin realizar ninguna acción.")
        else:
            logging.info(f"El contacto {numero} envió un mensaje irrelevante.")

# Función para procesar un número
def procesar_numero(row):
    global total_enviados, total_fallidos, total_no_portables
    numero = str(row['Numero']).strip()
    
    if not validar_numero(numero):
        logging.error(f"Número inválido: {numero}")
        return {'Numero': numero, 'Estado Portabilidad': "Número inválido", 'Estado Envío': "No enviado"}

    logging.info(f"Verificando número: {numero}")

    # Verificar el estado de portabilidad
    estado_portabilidad = verificar_portabilidad(numero)
    logging.info(f"Estado de portabilidad: {estado_portabilidad}")

    if estado_portabilidad == "El número de celular no tiene una solicitud de portabilidad en curso.":
        mensaje = obtener_saludo() + "\n" + mensaje_promocion
        numero = "+57" + numero if not numero.startswith("+") else numero

        intentos = 0
        while intentos < MAX_REINTENTOS:
            if enviar_mensaje(numero, mensaje):
                total_enviados += 1  # Incrementar el contador de mensajes enviados
                break
            intentos += 1
            time.sleep(PAUSA_INCREMENTAL * intentos)  # Aumentar la pausa con cada intento

        if intentos == MAX_REINTENTOS:
            total_fallidos += 1  # Incrementar el contador de fallidos
            estado_envio = "No enviado después de varios intentos"
        else:
            estado_envio = "Enviado"
    else:
        estado_envio = "No enviado - Portabilidad en curso"
        total_no_portables += 1  # Incrementar el contador de no portables

    return {'Numero': numero, 'Estado Portabilidad': estado_portabilidad, 'Estado Envío': estado_envio}

# Procesar los números en lotes
for i in range(0, len(df), MENSAJES_POR_LOTE):
    lote = df.iloc[i:i + MENSAJES_POR_LOTE]

    for index, row in lote.iterrows():
        resultado = procesar_numero(row)
        df.at[index, 'Estado Portabilidad'] = resultado['Estado Portabilidad']
        df.at[index, 'Estado Envío'] = resultado['Estado Envío']

    logging.info(f"Lote {i // MENSAJES_POR_LOTE + 1} procesado.")
    df.to_excel('resultados_portabilidad.xlsx', index=False)  # Guardar los resultados intermedios en Excel

    if i + MENSAJES_POR_LOTE < len(df):
        logging.info(f"Esperando {TIEMPO_ESPERA_LOTE // 60} minutos antes del siguiente lote...")
        mostrar_temporizador(TIEMPO_ESPERA_LOTE)
        analizar_mensajes_entrantes()  # Analizar mensajes entrantes durante la espera

<<<<<<< HEAD
# Guardar los resultados en un nuevo archivo Excel
df_resultados = pd.DataFrame(resultados)
df_resultados.to_excel('resultados_finales_portabilidad_y_mensajes_lote.xlsx', index=False)

# Resumen final del proceso
logging.info(f"Proceso completado: {total_enviados} mensajes enviados con éxito, {total_fallidos} fallos en el envío, {total_no_portables} números no portables.")

print(f"Resumen: {total_enviados} mensajes enviados con éxito, {total_fallidos} fallos en el envío, {total_no_portables} números no portables.")

=======
logging.info(f"Proceso completo: {total_enviados} enviados, {total_fallidos} fallidos, {total_no_portables} no portables.")
>>>>>>> cc979af3a2a3b0c5b0f90338727c387991250c2c
