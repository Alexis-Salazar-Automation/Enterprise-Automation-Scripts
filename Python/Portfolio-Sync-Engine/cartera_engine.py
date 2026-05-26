import os
import json
import requests
import pandas as pd
import gspread
import smtplib
from datetime import datetime
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
from gspread_dataframe import set_with_dataframe
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

# --- CONFIGURACIÓN DE ENTORNO ---
# Estas variables se configuran en los Secrets de GitHub
EMAIL_USER = os.environ.get("EMAIL_USER")
EMAIL_PASSWORD = os.environ.get("EMAIL_PASSWORD")
GOOGLE_SHEET_ID = os.environ.get("GOOGLE_SHEET_ID")
DRIVE_FOLDER_ID = os.environ.get("DRIVE_BACKUP_FOLDER_ID")
API_LOGIN_PAYLOAD = os.environ.get("API_LOGIN_PAYLOAD")

def enviar_correo(asunto, cuerpo_html, destinatarios):
    msg = MIMEMultipart("alternative")
    msg['Subject'] = asunto
    msg['From'] = EMAIL_USER
    msg['To'] = ", ".join(destinatarios)
    msg.attach(MIMEText(cuerpo_html, "html"))
    
    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(EMAIL_USER, EMAIL_PASSWORD)
            server.sendmail(EMAIL_USER, destinatarios, msg.as_string())
        print("✅ Reporte enviado por correo.")
    except Exception as e:
        print(f"❌ Error SMTP: {e}")

def ejecutar_bot():
    print("🚀 INICIANDO AUTOMATION CARTERA ENGINE...")
    inicio = datetime.now()
    
    # Configuración de reintentos
    session = requests.Session()
    retry = Retry(total=5, backoff_factor=2, status_forcelist=[500, 502, 503, 504])
    session.mount('https://', HTTPAdapter(max_retries=retry))

    try:
        # FASE 1: AUTENTICACIÓN
        print("🔑 Autenticando con la API...")
        url_login = "[API_ENDPOINT]"
        if not API_LOGIN_PAYLOAD:
            raise Exception("Falta Secreto: API_LOGIN_PAYLOAD")
            
        payload_login = json.loads(API_LOGIN_PAYLOAD)
        resp_login = session.post(url_login, json=payload_login, timeout=15)
        
        if resp_login.status_code != 200 or resp_login.json().get("status") != 1:
            raise Exception("Error en Login API.")
        
        mi_token = resp_login.json()["data"]["token"]

        # FASE 2: DESCARGA DE DATOS
        print("📥 Descargando cartera completa...")
        url_cartera = "[API_ENDPOINT]"
        headers = {"authorization": mi_token, "Content-Type": "application/json"}
        resp_cartera = session.post(url_cartera, headers=headers, json={"lotes": False, "get_ctes": True}, timeout=60)
        datos_cartera = resp_cartera.json()
        
        # Identificar la lista de datos en el JSON
        llave_datos = max((k for k in datos_cartera if isinstance(datos_cartera[k], list)), key=lambda k: len(datos_cartera[k]))
        df_completo = pd.DataFrame(datos_cartera[llave_datos])

        # FASE 3: TRANSFORMACIÓN (PANDAS)
        print("🧹 Mapeando columnas y limpiando datos...")
        cols_input = [COLUMNS_NAME]
        cols_output = [COLUMNS_RENAME]
        
        df_final = df_completo[cols_input].copy()
        df_final.fillna("", inplace=True) 
        df_final.columns = cols_output

        # FASE 4: COMPARACIÓN Y GOOGLE SHEETS
        print("☁️ Conectando a Google Sheets...")
        credenciales_json = json.loads(os.environ["CREDENCIALES_GOOGLE"])
        creds = service_account.Credentials.from_service_account_info(credenciales_json, 
                scopes=['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive'])
        gc = gspread.authorize(creds)
        
        libro = gc.open_by_key(GOOGLE_SHEET_ID)
        hoja_cartera = libro.worksheet("Cartera")
        
        # Obtener datos previos para comparar
        datos_viejos = hoja_cartera.get_all_values()
        referencias_nuevas = set(df_final["REFERENCIA"].astype(str))
        referencias_viejas = set()
        dict_viejos = {}

        if len(datos_viejos) > 1:
            for fila in datos_viejos[1:]:
                ref = str(fila[2]) # Índice de REFERENCIA
                referencias_viejas.add(ref)
                dict_viejos[ref] = fila[0] # Guardar nombre

        casos_eliminados = referencias_viejas - referencias_nuevas
        casos_nuevos = referencias_nuevas - referencias_viejas

        print("⌛ Actualizando hoja de cálculo...")
        hoja_cartera.batch_clear(['A2:Z100000'])
        set_with_dataframe(hoja_cartera, df_final, row=2, include_column_header=False)

        # FASE 5: RESPALDO EN DRIVE
        print("📁 Generando respaldo en Excel...")
        nombre_respaldo = f"Backup_Cartera_{datetime.now().strftime('%Y-%m-%d')}.xlsx"
        df_final.to_excel(nombre_respaldo, index=False)
        
        drive_service = build('drive', 'v3', credentials=creds)
        metadata = {'name': nombre_respaldo, 'parents': [DRIVE_FOLDER_ID]}
        media = MediaFileUpload(nombre_respaldo, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        
        archivo_drive = drive_service.files().create(body=metadata, media_body=media, fields='webViewLink').execute()
        link_drive = archivo_drive.get('webViewLink')

        # FASE 6: REPORTE POR CORREO
        print("✉️ Enviando reporte...")
        html_del = "".join([f"<li>🔴 {dict_viejos.get(r, 'N/A')} (Ref: {r})</li>" for r in list(casos_eliminados)[:15]])
        html_new = "".join([f"<li>🟢 Referencia nueva: {r}</li>" for r in list(casos_nuevos)[:15]])

        cuerpo = f"""
        <html>
            <body style="font-family: sans-serif;">
                <h2 style="color: #0A3D62;">📊 Resumen de Actualización</h2>
                <p>La cartera se ha sincronizado correctamente.</p>
                <ul>
                    <li><b>Total registros:</b> {len(df_final)}</li>
                    <li><b>Casos nuevos:</b> {len(casos_nuevos)}</li>
                    <li><b>Casos removidos:</b> {len(casos_eliminados)}</li>
                </ul>
                <a href="{link_drive}" style="background: #2980B9; color: white; padding: 10px; text-decoration: none; border-radius: 5px;">Descargar Excel de Respaldo</a>
                <br><br>
                <h4>Cambios detectados:</h4>
                <ul>{html_del if html_del else "<li>✅ Sin bajas</li>"}</ul>
                <ul>{html_new if html_new else "<li>➖ Sin altas</li>"}</ul>
            </body>
        </html>
        """
        enviar_correo("✅ Cartera Actualizada", cuerpo, [EMAIL_USER])
        print("✅ PROCESO COMPLETADO.")

    except Exception as e:
        print(f"❌ ERROR: {e}")
        error_html = f"<h3>🚨 Error en el Bot</h3><p>{str(e)}</p>"
        enviar_correo("🚨 Fallo en Cartera Engine", error_html, [EMAIL_USER])
        raise e

if __name__ == "__main__":
    ejecutar_bot()
