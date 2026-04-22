import win32com.client
import pythoncom
import time
import os
import re
import threading
import requests
import sys
import calendar
from datetime import datetime
import gc
import win32event
import win32api
import winerror

# Configuracion de instancia unica (Mutex)
mutex = win32event.CreateMutex(None, False, "Global\\REPORTE_NISSAN_MUTEX")
if win32api.GetLastError() == winerror.ERROR_ALREADY_EXISTS:
    sys.exit(0)

print("SCRIPT ARRANCÓ")

DIA_LIMITE_SOLICITADO = int(sys.argv[1]) if len(sys.argv) > 1 else 31
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
                print("Sucursal:", SUCURSAL_BUSQUEDA)
                return True
        return False
    except Exception as e:
        print("Error sucursal:", e)
        return False

def convertir_fecha(valor):
    if not valor:
        return None
    if isinstance(valor, datetime):
        return valor
    match = re.search(r"(\d{1,2})[-/](\d{1,2})", str(valor))
    if match:
        try:
            return datetime(2026, int(match.group(2)), int(match.group(1)))
        except:
            return None
    return None

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
        url = f"{API_URL_GLOBALS}?filters[fecha][$eq]={fecha_iso}&filters[sucursal][documentId][$eq]={DOC_ID_SUCURSAL}"
        res_get = session.get(url)
        if not res_get.ok:
            return
        data = res_get.json().get("data", [])
        if not data:
            session.post(API_URL_GLOBALS, json=payload)
            return
        actual = data[0]
        doc_id = actual["documentId"]
        cambios = False
        for campo, nuevo_valor in payload["data"].items():
            if campo not in actual:
                continue
            actual_valor = actual[campo]
            if isinstance(nuevo_valor, (int, float)):
                if nuevo_valor < actual_valor:
                    payload["data"][campo] = actual_valor
                if nuevo_valor > actual_valor:
                    cambios = True
            else:
                if nuevo_valor != actual_valor:
                    cambios = True
        if cambios:
            session.put(f"{API_URL_GLOBALS}/{doc_id}", json=payload)
    except Exception as e:
        print("Error guardando:", e)

def extraer_bloque_posicional(ws, dia_inicio, dia_limite):
    data_range = ws.Range(ws.Cells(6, 1), ws.Cells(12, 65)).Value
    dia_fin = dia_limite if dia_inicio >= 22 else min(dia_inicio + 6, dia_limite)
    for fila in data_range:
        fecha = convertir_fecha(fila[8])
        if fecha and dia_inicio <= fecha.day <= dia_fin:
            def leer(col):
                try:
                    return int(float(fila[col-1])) if fila[col-1] else 0
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

def ejecutar():
    print("Entrando ejecutar")
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

        # TUS MACROS INTACTAS
        try:
            excel.Application.Run("MostrarRangoCALENDARIO")
            excel.Application.Run("CambiarColorAzul_Boton6")
            excel.Application.Run("CambiarColorAzul_Boton1")
            time.sleep(2)
        except:
            pass

        # Obtenemos mes y año de la plantilla (usamos la fecha de la celda de datos si la 2,185 falla)
        # O por defecto el mes actual si no hay referencia
        hoy = datetime.now()
        
        # Intentamos obtener la fecha de la plantilla para saber en qué mes estamos
        try:
            # Buscamos en la fila 6 col 9 (donde empieza el reporte) para saber el mes
            fecha_ref = convertir_fecha(ws.Cells(6, 9).Value)
            anio = fecha_ref.year if fecha_ref else hoy.year
            mes = fecha_ref.month if fecha_ref else hoy.month
        except:
            anio, mes = hoy.year, hoy.month

        # Restricción: Si es el mes actual, solo permitir hasta el día anterior
        DIA_LIMITE = DIA_LIMITE_SOLICITADO
        if anio == hoy.year and mes == hoy.month:
            if DIA_LIMITE >= hoy.day:
                DIA_LIMITE = hoy.day - 1
        
        if DIA_LIMITE < 1:
            print("Día solicitado no procesable todavía")
            return

        # CALCULO DE LA POSICION DEL DIA 1
        # weekday() devuelve 0 para Lunes, 6 para Domingo. 
        # En calendarios estándar (Dom-Sab), el offset es: (Lunes=1, ..., Sab=6, Dom=0)
        primer_dia_mes = datetime(anio, mes, 1)
        offset = (primer_dia_mes.weekday() + 1) % 7 

        bloques = [1, 8, 15, 22]
        if DIA_LIMITE >= 29:
            bloques.append(29)

        for i, dia in enumerate(bloques):
            if dia > DIA_LIMITE:
                continue

            # El calculo de tabs: 4 base + offset del día 1 + saltos de semana
            tabs_semana = offset + (i * 7)

            def hilo():
                pythoncom.CoInitialize()
                try:
                    shell = win32com.client.Dispatch("WScript.Shell")
                    time.sleep(2.5)
                    if shell.AppActivate("CALENDARIO") or shell.AppActivate("Microsoft Excel"):
                        shell.SendKeys("{TAB 4}") # Posicion 1 del calendario
                        if tabs_semana > 0:
                            shell.SendKeys(f"{{TAB {tabs_semana}}}")
                        shell.SendKeys("{ENTER}")
                        time.sleep(1)
                        shell.SendKeys("%{F4}")
                finally:
                    pythoncom.CoUninitialize()

            t = threading.Thread(target=hilo)
            t.start()

            try:
                excel.Application.Run("MostrarCalendario")
            except:
                pass

            t.join()
            time.sleep(2)
            extraer_bloque_posicional(ws, dia, DIA_LIMITE)

        wb.Save()
        print("PROCESO TERMINADO EXITOSAMENTE")

    finally:
        try:
            wb.Close(False)
        except:
            pass
        excel.Quit()
        gc.collect()

if __name__ == "__main__":
    ejecutar()