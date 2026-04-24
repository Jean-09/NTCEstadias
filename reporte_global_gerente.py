import win32com.client
import pythoncom
import time
import os
import re
import threading
import requests
import sys
from datetime import datetime
import gc
import win32event
import win32api
import winerror
import pyautogui
import pytesseract
import cv2
import numpy as np

# Configuración OCR
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

# Control de instancia única
mutex = win32event.CreateMutex(None, False, "Global\\REPORTE_GERENTE_NISSAN_MUTEX")
if win32api.GetLastError() == winerror.ERROR_ALREADY_EXISTS:
    sys.exit(0)

# Parámetros
DIA_LIMITE = int(sys.argv[1]) if len(sys.argv) > 1 else 31
TOKEN = sys.argv[2] if len(sys.argv) > 2 else None
SUCURSAL_BUSQUEDA = (sys.argv[3] if len(sys.argv) > 3 else "Zacatecas").strip()

API_BASE = "http://localhost:1337/api"
API_URL_GLOBALS = f"{API_BASE}/globals"
API_URL_SUCURSALES = f"{API_BASE}/sucursals"

session = requests.Session()
if TOKEN:
    session.headers.update({"Authorization": f"{TOKEN}"})

DOC_ID_SUCURSAL = None
NOMBRE_ARCHIVO = None

def obtener_configuracion_sucursal():
    global DOC_ID_SUCURSAL, NOMBRE_ARCHIVO
    try:
        res = session.get(f"{API_URL_SUCURSALES}?filters[Sucursal][$eq]={SUCURSAL_BUSQUEDA}")
        if res.ok:
            datos = res.json().get("data", [])
            if datos:
                suc = datos[0]
                DOC_ID_SUCURSAL = suc.get("documentId")
                archivo_db = suc.get("Plantilla")
                user_downloads = os.path.join(os.path.expanduser("~"), "Downloads")
                NOMBRE_ARCHIVO = os.path.join(user_downloads, archivo_db)
                return True
        return False
    except: return False

def convertir_fecha(valor):
    if not valor: return None
    if isinstance(valor, datetime): return valor
    match = re.search(r"(\d{1,2})[-/](\d{1,2})", str(valor))
    if match:
        try: return datetime(2026, int(match.group(2)), int(match.group(1)))
        except: return None
    return None

def guardar_en_strapi(reg, nombre_gerente):
    fecha_iso = reg["OBJ_FECHA"].strftime("%Y-%m-%d")
    payload = {
        "data": {
            "preContactos": reg["PRE_CONTACTOS"],
            "Contactos": reg["CONTACTOS"],
            "prospectos": reg["PROSPECTOS"],
            "solCDatosCompletos": reg["SOL_DATOS_COMPLETOS"],
            "viablesPreAutorizadas": reg["VIABLES"],
            "citasAgendadas": reg["CITAS_AGENDADAS"],
            "citasReales": reg["CITAS_REALES"],
            "docCompleta": reg["DOC_COMPLETA"],
            "autorizadas": reg["AUTORIZADAS"],
            "pedidosConAnticipo": reg["PEDIDO_ANTICIPO"],
            "demos": reg["DEMOS"],
            "entregas": reg["ENTREGAS"],
            "desembolsos": reg["DESEMBOLSOS"],
            "fecha": fecha_iso,
            "sucursal": DOC_ID_SUCURSAL,
            "tipo": "Gerente",
            "Gerente": nombre_gerente
        }
    }
    try:
        # Buscar si ya existe el registro para evitar duplicados
        url_busqueda = f"{API_URL_GLOBALS}?filters[fecha][$eq]={fecha_iso}&filters[sucursal][documentId][$eq]={DOC_ID_SUCURSAL}&filters[Gerente][$eq]={nombre_gerente}"
        res_get = session.get(url_busqueda)
        existente = res_get.json().get("data", [])
        
        if existente:
            doc_id = existente[0].get("documentId")
            session.put(f"{API_URL_GLOBALS}/{doc_id}", json=payload)
        else:
            session.post(API_URL_GLOBALS, json=payload)
    except Exception as e:
        print(f"! Error al guardar en Strapi: {e}")

def obtener_nombre_gerente_ocr(indice, contenedor):
    time.sleep(2.0)
    try:
        if indice == 0:
            pyautogui.press("up", presses=100)
        else:
            pyautogui.press("down")

        time.sleep(1.2)
        screenshot = pyautogui.screenshot()
        img = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)
        
        mask = cv2.inRange(img, np.array([100, 40, 0]), np.array([255, 160, 60]))
        contornos, _ = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

        for c in contornos:
            x, y, w, h = cv2.boundingRect(c)
            if 100 < w < 600 and 15 < h < 60:
                recorte = cv2.cvtColor(img[y:y+h, x:x+w], cv2.COLOR_BGR2GRAY)
                _, thresh = cv2.threshold(recorte, 150, 255, cv2.THRESH_BINARY_INV)
                
                nombre_raw = pytesseract.image_to_string(thresh, config="--psm 7").strip()
                nombre_limpio = re.sub(r"[^a-zA-Z\s]", "", nombre_raw).strip().upper()
                
                if len(nombre_limpio) > 3 and nombre_limpio not in ["LEADS", "CDC", "PV", "PISO"]:
                    contenedor["nombre"] = nombre_limpio
                    print(f"> Gerente Detectado: {nombre_limpio}") 
                    break
        pyautogui.hotkey("alt", "f4")
    except Exception as e:
        print(f"! Error OCR: {e}")

def hilo_calendario(pasos):
    time.sleep(1.5)
    pyautogui.press("tab", presses=4)
    if pasos > 0: pyautogui.press("tab", presses=pasos)
    pyautogui.press("enter")
    time.sleep(0.5)
    pyautogui.hotkey("alt", "f4")

def extraer_validar(ws, dias_procesados, nombre_gerente):
    try:
        data_range = ws.Range(ws.Cells(6, 1), ws.Cells(12, 65)).Value
        if not data_range: return 0
        max_dia = 0
        for fila in data_range:
            fecha = convertir_fecha(fila[8])
            if fecha:
                dia = fecha.day
                if dia > max_dia: max_dia = dia
                if dia not in dias_procesados and dia <= DIA_LIMITE:
                    def leer(col):
                        try:
                            val = fila[col-1]
                            return int(float(val)) if val is not None else 0
                        except: return 0
                    
                    reg = {
                        "OBJ_FECHA": fecha,
                        "PRE_CONTACTOS": leer(12), "CONTACTOS": leer(13), "PROSPECTOS": leer(14),
                        "SOL_DATOS_COMPLETOS": leer(47), "VIABLES": leer(51),
                        "CITAS_AGENDADAS": leer(15), "CITAS_REALES": leer(16),
                        "DOC_COMPLETA": leer(53), "AUTORIZADAS": leer(56),
                        "PEDIDO_ANTICIPO": leer(61), "DEMOS": leer(45),
                        "ENTREGAS": leer(65), "DESEMBOLSOS": leer(63),
                    }
                    
                    guardar_en_strapi(reg, nombre_gerente)
                    dias_procesados.add(dia)
        return max_dia
    except Exception as e:
        print(f"! Error en extracción: {e}")
        return 0

def ejecutar():
    print(f"--- Iniciando: {SUCURSAL_BUSQUEDA} ---")
    if not obtener_configuracion_sucursal() or not os.path.exists(NOMBRE_ARCHIVO):
        print("! Error: No se encontró el archivo de Excel.")
        return
    
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    try:
        wb = excel.Workbooks.Open(NOMBRE_ARCHIVO)
        ws = wb.Sheets("Reportes")
        
        try:
            excel.Application.Run("MostrarRangoCALENDARIO")
            excel.Application.Run("CambiarColorAzul_Boton6")
            excel.Application.Run("Modulo2.CambiarColorAzul_Boton13")
            time.sleep(1.2)
        except: pass

        raw_f = str(ws.Cells(2, 185).Value).replace("*", "").strip()
        f_base = datetime.strptime(raw_f, "%d-%m-%Y")
        offset = (datetime(f_base.year, f_base.month, 1).weekday() + 1) % 7

        gerentes_vistos = []
        idx_g = 0
        fallos_consecutivos = 0
        
        while True:
            contenedor = {"nombre": None}
            t_ocr = threading.Thread(target=obtener_nombre_gerente_ocr, args=(idx_g, contenedor))
            t_ocr.start()
            try: excel.Application.Run("CambiarColorAzul_Boton3")
            except: pass
            t_ocr.join()

            nombre = contenedor["nombre"]
            
            if not nombre or (gerentes_vistos and nombre == gerentes_vistos[-1]):
                fallos_consecutivos += 1
                if fallos_consecutivos >= 2: break
                idx_g += 1
                continue

            fallos_consecutivos = 0
            gerentes_vistos.append(nombre)
            dias_procesados = set()
            ultimo_dia = 0
            vuelta = 0

            time.sleep(1.5)

            while ultimo_dia < DIA_LIMITE and vuelta < 6:
                t_cal = threading.Thread(target=hilo_calendario, args=(offset + (vuelta * 7),))
                t_cal.start()
                try: excel.Application.Run("MostrarCalendario")
                except: pass
                t_cal.join()

                time.sleep(3.8) 
                max_leido = extraer_validar(ws, dias_procesados, nombre)
                
                if max_leido == ultimo_dia and max_leido < DIA_LIMITE:
                    time.sleep(2.0)
                    max_leido = extraer_validar(ws, dias_procesados, nombre)
                
                ultimo_dia = max_leido
                vuelta += 1

            idx_g += 1
            gc.collect()
            
    finally:
        try: wb.Close(False)
        except: pass
        excel.Quit()
        print("--- Proceso Terminado ---")
        gc.collect()

if __name__ == "__main__":
    try: ejecutar()
    finally: win32api.CloseHandle(mutex)