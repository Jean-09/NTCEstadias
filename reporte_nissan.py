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

# Configuracion de instancia unica (Mutex)
mutex = win32event.CreateMutex(None, False, "Global\\REPORTE_NISSAN_MUTEX")
if win32api.GetLastError() == winerror.ERROR_ALREADY_EXISTS:
    sys.exit(0)

# Argumentos de linea de comandos
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

# Obtiene datos de la sucursal desde Strapi
def obtener_configuracion_sucursal():
    global DOC_ID_SUCURSAL, NOMBRE_ARCHIVO
    try:
        res = session.get(
            f"{API_URL_SUCURSALES}?filters[Sucursal][$eq]={SUCURSAL_BUSQUEDA}"
        )
        if res.ok:
            datos = res.json().get("data", [])
            if datos:
                suc = datos[0]
                DOC_ID_SUCURSAL = suc.get("documentId")
                archivo_db = suc.get("Plantilla")
                user_downloads = os.path.join(os.path.expanduser("~"), "Downloads")
                NOMBRE_ARCHIVO = os.path.join(user_downloads, archivo_db)
                print(f"Sucursal: {SUCURSAL_BUSQUEDA} (ID: {DOC_ID_SUCURSAL})")
                return True
        return False
    except Exception:
        return False

# Procesa formatos de fecha desde Excel
def convertir_fecha(valor):
    if not valor:
        return None
    if isinstance(valor, datetime):
        return valor
    texto = str(valor).strip()
    match = re.search(r"(\d{1,2})[-/](\d{1,2})", texto)
    if match:
        try:
            dia, mes = int(match.group(1)), int(match.group(2))
            return datetime(2026, mes, dia)
        except:
            return None
    return None

# Envia o actualiza los datos en la API de Strapi
def guardar_en_strapi(reg):
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
            "tipo": "Global",
        }
    }

    try:
        check_url = f"{API_URL_GLOBALS}?filters[fecha][$eq]={fecha_iso}&filters[sucursal][documentId][$eq]={DOC_ID_SUCURSAL}"
        res_get = session.get(check_url)
        existente = res_get.json().get("data", [])

        if existente:
            doc_id = existente[0].get("documentId")
            res = session.put(f"{API_URL_GLOBALS}/{doc_id}", json=payload)
        else:
            res = session.post(API_URL_GLOBALS, json=payload)

        if not res.ok:
            print(f"Error {res.status_code}: {res.text}")
        else:
            print(f"Procesado: {fecha_iso}")
    except Exception as e:
        print(f"Error: {e}")

# Lee los datos de las celdas de Excel y valida si encontro el dia objetivo
def extraer_bloque_posicional(ws, dia_inicio, dia_limite):
    data_range = ws.Range(ws.Cells(6, 1), ws.Cells(12, 65)).Value
    dia_fin = dia_limite if dia_inicio >= 22 else min(dia_inicio + 6, dia_limite)
    encontrado_objetivo = False
    encontrados_total = 0

    for i in range(len(data_range)):
        fila_data = data_range[i]
        fecha = convertir_fecha(fila_data[8])
        if fecha and dia_inicio <= fecha.day <= dia_fin:
            if fecha.day == dia_inicio:
                encontrado_objetivo = True
            
            def leer(col):
                v = fila_data[col - 1]
                try:
                    return int(float(v)) if v is not None else 0
                except:
                    return 0

            registro = {
                "OBJ_FECHA": fecha,
                "PRE_CONTACTOS": leer(12),
                "CONTACTOS": leer(13),
                "PROSPECTOS": leer(14),
                "SOL_DATOS_COMPLETOS": leer(47),
                "VIABLES": leer(51),
                "CITAS_AGENDADAS": leer(15),
                "CITAS_REALES": leer(16),
                "DOC_COMPLETA": leer(53),
                "AUTORIZADAS": leer(56),
                "PEDIDO_ANTICIPO": leer(61),
                "DEMOS": leer(45),
                "ENTREGAS": leer(65),
                "DESEMBOLSOS": leer(63),
            }
            guardar_en_strapi(registro)
            encontrados_total += 1

    print(f"Bloque {dia_inicio}: {encontrados_total} procesados.")
    return encontrado_objetivo

# Automatiza la navegacion del calendario mediante teclado
def hilo_teclado(dia, ajuste_tabs=0):
    pythoncom.CoInitialize()
    try:
        shell = win32com.client.Dispatch("WScript.Shell")
        time.sleep(2.5)
        if shell.AppActivate("CALENDARIO") or shell.AppActivate("Microsoft Excel"):
            pestañas = {1: 0, 8: 7, 15: 14, 22: 21, 29: 28}
            saltos = pestañas.get(dia, 0) + ajuste_tabs
            shell.SendKeys("{TAB 4}")
            if saltos > 0:
                shell.SendKeys(f"{{TAB {saltos}}}")
            shell.SendKeys("{ENTER}")
            time.sleep(1.0)
            shell.SendKeys("%{F4}")
    finally:
        pythoncom.CoUninitialize()

# Orquestador principal del proceso
def ejecutar():
    if not obtener_configuracion_sucursal():
        return
    if not os.path.exists(NOMBRE_ARCHIVO):
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
            excel.Application.Run("CambiarColorAzul_Boton1")
            time.sleep(2)
        except:
            pass

        puntos_inicio = [1, 8, 15, 22]
        if DIA_LIMITE >= 29:
            puntos_inicio.append(29)

        for inicio in puntos_inicio:
            if inicio > DIA_LIMITE:
                continue
            
            exito = False
            intentos_tabs = 0
            
            # Intenta hasta encontrar el dia correcto ajustando tabs si es necesario
            while not exito and intentos_tabs < 8:
                print(f"--- Intentando Bloque dia {inicio} (Ajuste TAB: {intentos_tabs}) ---")
                t = threading.Thread(target=hilo_teclado, args=(inicio, intentos_tabs))
                t.start()
                try:
                    excel.Application.Run("MostrarCalendario")
                except:
                    pass
                t.join()
                time.sleep(2.0)
                
                exito = extraer_bloque_posicional(ws, inicio, DIA_LIMITE)
                intentos_tabs += 1
                
        wb.Save()
    finally:
        try:
            wb.Close(False)
        except:
            pass
        excel.Quit()
        gc.collect()

if __name__ == "__main__":
    ejecutar()