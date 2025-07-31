from components.consumer_api import getdata
from components.create_excel import create_file
from components.outlook import abrir_outlook
from components.email import conectar_a_outlook, enviar_correo_con_adjunto,obtener_cuenta_outlook
from utils.config import TORRES, DESTINATARIOS, CC_DESTINATARIOS
from utils.email_body import cuerpo_post
from dotenv import load_dotenv
from pathlib import Path
import os

# Cargar variables desde el archivo .env
load_dotenv()

# CARPETAS
excel_folder = Path(os.path.abspath(r".\assets"))
excel_folder.mkdir(parents=True, exist_ok=True)

# URL de la API
API_LINEAS_NO_PROVISIONADAS = os.getenv("API_LINEAS_NO_PROVISIONADAS")

# Obtener datos de la API
dataNoProvisionada = getdata(API_LINEAS_NO_PROVISIONADAS)

# # Guardar en un solo archivo con varias hojas
excel_path = os.path.join(excel_folder, "Linea no provisionada.xlsx")

create_file(dataNoProvisionada['data'], excel_path)

# ENVIO CORREO ELECTRONICO
abrir_outlook()
outlook = conectar_a_outlook()
obtener_cuenta_outlook(outlook,"reportesbi.soap@hitss.com")

destinatario = DESTINATARIOS.get(TORRES[0])
print(destinatario)
asunto = f"Alineacion de su SDP lineas ONE."
cuerpo = cuerpo_post()

enviar_correo_con_adjunto(destinatario, asunto, cuerpo, excel_path, CC_DESTINATARIOS)