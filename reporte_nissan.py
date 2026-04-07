from csv import excel

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

# -------- BLOQUEO PARA EVITAR DOBLE EJECUCIÓN --------
mutex = win32event.CreateMutex(None, False, "REPORTE_NISSAN_MUTEX")

if win32api.GetLastError() == winerror.ERROR_ALREADY_EXISTS:
    print("Ya hay una instancia ejecutándose. Cancelado.")
    sys.exit(0)
    
# --- CONFIGURACIÓN DINÁMICA ---
DIA_LIMITE = int(sys.argv[1]) if len(sys.argv) > 1 else 31
TOKEN = sys.argv[2] if len(sys.argv) > 2 else None
SUCURSAL = (sys.argv[3] if len(sys.argv) > 3 else "Zacatecas").strip()

if SUCURSAL.lower() == "bicentenario":
    NOMBRE_ARCHIVO = r"C:\Users\jeanl\Downloads\jgyn5k0aZuBmDgkvmusCNA8h.xlsm"
else:
    NOMBRE_ARCHIVO = r"C:\Users\jeanl\Downloads\4mYaHHSktW5igUO5zi6FEYU6.xlsm"

HOJA_DATOS = "Reportes"
MACRO_CALENDARIO = "MostrarCalendario"
API_URL = "http://localhost:1337/api/globals"

session = requests.Session()
session.headers.update({"Authorization": TOKEN} if TOKEN else {})

cache = {}

# ---------------- STRAPI ----------------

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
            "Sucursal": SUCURSAL
        }
    }

    print("\n" + "="*50)
    print(f"PROCESANDO FECHA: {fecha_iso}")
    print(f"SUCURSAL: {SUCURSAL}")
    print(f"DATOS: {payload['data']}")
    print("="*50)

    try:
        res = session.post(API_URL, json={"data": payload["data"]})

        print(f"STATUS: {res.status_code}")

        if res.status_code in [200, 201]:
            print(f"[OK] GUARDADO {fecha_iso}")
        else:
            print(f"[ERROR] {fecha_iso}")
            print("RESPUESTA STRAPI:")
            print(res.text)

    except Exception as e:
        print(f"[ERROR CONEXIÓN] {fecha_iso}")
        print(e)
# ---------------- EXCEL ----------------

def extraer_bloque_posicional(ws, dia_inicio, dia_limite):
    ultima_fila = ws.UsedRange.Rows.Count

    if dia_inicio == 22:
        dia_fin = dia_limite
    else:
        dia_fin = min(dia_inicio + 6, dia_limite)

    fechas_procesadas = set()

    for fila in range(2, ultima_fila + 1):

        v9 = ws.Cells(fila, 9).Value
        v10 = ws.Cells(fila, 10).Value

        fecha = None

        fecha = convertir_fecha(v9) or convertir_fecha(v10)

        if not fecha or not (dia_inicio <= fecha.day <= dia_fin):
            continue

        f_id = fecha.strftime("%Y-%m-%d")
        if f_id in fechas_procesadas:
            continue

        fechas_procesadas.add(f_id)

        def val(p):
            try:
                v = ws.Cells(fila, p).Value
                return int(float(v)) if v else 0
            except:
                return 0

        registro = {
            "OBJ_FECHA": fecha,
            "PRE_CONTACTOS": val(12),
            "CONTACTOS": val(13),
            "PROSPECTOS": val(14),
            "SOL_DATOS_COMPLETOS": val(47),
            "VIABLES": val(51),
            "CITAS_AGENDADAS": val(15),
            "CITAS_REALES": val(16),
            "DOC_COMPLETA": val(53),
            "AUTORIZADAS": val(55),
            "PEDIDO_ANTICIPO": val(61),
            "DEMOS": val(45),
            "ENTREGAS": val(65),
            "DESEMBOLSOS": val(63)
        }
        print(f"Extrayendo fila {fila} | Fecha detectada: {fecha}")

        guardar_en_strapi(registro)

# ---------------- MAIN ----------------

def ejecutar():

    if not os.path.exists(NOMBRE_ARCHIVO):
        print("No existe el archivo.")
        return

    excel = None

    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.AutomationSecurity = 1
        excel.ScreenUpdating = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        excel.Visible = True

        wb = excel.Workbooks.Open(NOMBRE_ARCHIVO)
        ws = wb.Sheets(HOJA_DATOS)

        bloques = [b for b in [1, 8, 15, 22] if b <= DIA_LIMITE]

        if DIA_LIMITE > 28:
            bloques.append(29)

        for inicio in bloques:

            t = threading.Thread(target=hilo_teclado, args=(inicio,))
            t.start()

            try:
                excel.Application.Run(f"'{wb.Name}'!{MACRO_CALENDARIO}")
            except Exception as e:
                print(f"Error macro: {e}")

            t.join()
            time.sleep(1.5)

            extraer_bloque_posicional(ws, inicio, DIA_LIMITE)

        wb.Save()
        print("FIN PROCESO")

    finally:
        if excel:
            try:
                wb.Close(False)
            except:
                pass

            try:
                excel.Quit()
            except:
                pass

            excel.DisplayAlerts = False
            wb.Close(SaveChanges=False)
            excel.Quit()
            time.sleep(1)
            gc.collect()
            try:
                win32api.CloseHandle(mutex)
            except:
                pass

# ---------------- UTILS ----------------

def convertir_fecha(valor):
    if not valor:
        return None

    texto = str(valor).strip()
    match = re.search(r'(\d{1,2})[-/](\d{1,2})', texto)

    if match:
        dia, mes = int(match.group(1)), int(match.group(2))

        if mes == datetime.now().month:
            return datetime(datetime.now().year, mes, dia)

    return None

def hilo_teclado(dia):
    pythoncom.CoInitialize()

    try:
        shell = win32com.client.Dispatch("WScript.Shell")

        time.sleep(2)

        for _ in range(5):
            if shell.AppActivate("Microsoft Excel"):
                break
            time.sleep(0.5)

        time.sleep(0.5)

        # 🔥 POSICIÓN BASE CORRECTA
        shell.SendKeys("{TAB 4}")
        time.sleep(0.3)

        # 🔥 AJUSTE POR BLOQUE
        if dia == 8:
            shell.SendKeys("{TAB 7}")
        elif dia == 15:
            shell.SendKeys("{TAB 14}")
        elif dia == 22:
            shell.SendKeys("{TAB 21}")
        elif dia == 29:
            shell.SendKeys("{TAB 28}")

        time.sleep(0.3)

        shell.SendKeys("{ENTER}")
        time.sleep(0.8)

        shell.SendKeys("%{F4}")

    finally:
        pythoncom.CoUninitialize()
# ---------------- RUN ----------------

if __name__ == "__main__":
    ejecutar()