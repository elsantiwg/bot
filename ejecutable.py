import pandas as pd
import requests
from bs4 import BeautifulSoup
import pywhatkit as pwk
import time
from datetime import datetime
import pytz

# Definir la URL de la página donde se realiza la consulta
url = "http://190.13.96.92/minisitio/Consulta_Numero_Celular.aspx"

# Cargar el archivo Excel con los números
file_path = 'numeros_portabilidad.xlsx'  # El archivo debe estar en la misma carpeta
df = pd.read_excel(file_path)

# Crear una sesión para mantener las cookies
session = requests.Session()

# Hacer una solicitud inicial para obtener los valores de VIEWSTATE y EVENTVALIDATION
response = session.get(url)
soup = BeautifulSoup(response.content, 'html.parser')

# Obtener los valores de VIEWSTATE, VIEWSTATEGENERATOR y EVENTVALIDATION necesarios
viewstate = soup.find("input", {"id": "__VIEWSTATE"}).get("value")
viewstate_generator = soup.find("input", {"id": "__VIEWSTATEGENERATOR"}).get("value")
event_validation = soup.find("input", {"id": "__EVENTVALIDATION"}).get("value")

# Parámetros configurables para el envío por lotes
mensajes_por_lote = 5          # Número de mensajes que se enviarán por cada lote
tiempo_espera_lote = 600       # Tiempo de espera entre lotes en segundos (10 minutos)
tiempo_espera_envio = 15       # Tiempo de espera para que cargue la página (ajústalo según la velocidad de carga)

# Obtener la hora actual en la zona horaria de Colombia
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
    response = session.post(url, data=payload)
    soup = BeautifulSoup(response.content, 'html.parser')
    # Extraer el mensaje de estado desde el HTML resultante
    mensaje_estado = soup.find("span", {"id": "lblInfoEstadoPortabilidad"}).text
    return mensaje_estado

# Función para mostrar el temporizador en segundos
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

# Iterar por cada número en el Excel y procesar en lotes
for i in range(0, len(df), mensajes_por_lote):
    # Obtener el lote actual de números
    lote = df.iloc[i:i + mensajes_por_lote]

    for index, row in lote.iterrows():
        numero = str(row['Numero'])  # Asegurarse de que el número sea un string
        print(f"Verificando número: {numero}")

        # Verificar el estado de portabilidad
        estado_portabilidad = verificar_portabilidad(numero)
        print(f"Estado de portabilidad: {estado_portabilidad}")

        # Si no tiene portabilidad en curso, se envía el mensaje
        if estado_portabilidad == "El número de celular no tiene una solicitud de portabilidad en curso.":
            # Obtener el saludo adecuado
            mensaje = obtener_saludo()

            # Agregar el código de país si no está presente
            if not numero.startswith("+"):
                numero = "+57" + numero  # Ajusta el código de país según sea necesario

            try:
                # Enviar el mensaje usando pywhatkit
                print(f"Enviando mensaje a {numero}: {mensaje}")
                pwk.sendwhatmsg_instantly(
                    numero, 
                    mensaje, 
                    tiempo_espera_envio,  # Esperar para cargar la página y enviar
                    True
                )  
                estado_envio = "Enviado"

                # Esperar unos segundos antes de enviar el siguiente mensaje en el mismo lote
                time.sleep(5)  # Puedes ajustar este tiempo si es necesario

            except Exception as e:
                print(f"Error al enviar mensaje a {numero}: {e}")
                estado_envio = f"Error: {e}"
        else:
            estado_envio = "No enviado - Portabilidad en curso"

        # Agregar los resultados al DataFrame
        resultados.append({'Numero': numero, 'Estado Portabilidad': estado_portabilidad, 'Estado Envío': estado_envio})

    # Esperar el tiempo configurado entre lotes, si no es el último lote
    if i + mensajes_por_lote < len(df):
        print(f"Esperando {tiempo_espera_lote / 60} minutos antes de enviar el siguiente lote...")
        mostrar_temporizador(tiempo_espera_lote)

# Guardar los resultados en un nuevo archivo Excel
df_resultados = pd.DataFrame(resultados)
df_resultados.to_excel('resultados_finales_portabilidad_y_mensajes_lote.xlsx', index=False)

print("Proceso de verificación y envío de mensajes completado.")