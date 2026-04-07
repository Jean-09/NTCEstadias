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

# --- CONFIGURACIÓN ---
mutex = win32event.CreateMutex(None, False, "Global\\REPORTE_NISSAN_MUTEX")
if win32api.GetLastError() == winerror.ERROR_ALREADY_EXISTS:
    print("Ya hay una instancia ejecutándose. Cancelado.")
    sys.exit(0)
    
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
# Mantenemos tu lógica de Token exacta
if TOKEN:
    session.headers.update({"Authorization": f"{TOKEN}"})

def convertir_fecha(valor):
    if not valor: return None
    if isinstance(valor, datetime): return valor
    texto = str(valor).strip()
    match = re.search(r'(\d{1,2})[-/](\d{1,2})', texto)
    if match:
        try:
            dia, mes = int(match.group(1)), int(match.group(2))
            return datetime(2026, mes, dia)
        except: return None
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
            "Sucursal": SUCURSAL,
            "tipo": "Global"
        }
    }

    try:
        res_get = session.get(f"{API_URL}?filters[fecha][$eq]={fecha_iso}&filters[Sucursal][$eq]={SUCURSAL}")
        if not res_get.ok:
            print(f" ERROR Strapi (GET {fecha_iso}): {res_get.status_code}")
            return

        existente = res_get.json().get('data', [])
        if existente:
            doc_id = existente[0].get('documentId') or existente[0].get('id')
            res = session.put(f"{API_URL}/{doc_id}", json=payload)
            action = "Actualizado"
        else:
            res = session.post(API_URL, json=payload)
            action = "Creado"
        
        if res.ok:
            print(f" ✅ {action}: {fecha_iso}")
        else:
            print(f" ❌ Error {action} {fecha_iso}: {res.status_code}")
    except Exception as e:
        print(f" Error de conexión ({fecha_iso}): {e}")

def extraer_bloque_posicional(ws, dia_inicio, dia_limite):
    # Lectura rápida de rango completo (Filas 6-12, Columnas 1-65)
    data_range = ws.Range(ws.Cells(6, 1), ws.Cells(12, 65)).Value
    dia_fin = dia_limite if dia_inicio >= 22 else min(dia_inicio + 6, dia_limite)
    encontrados = 0

    for i in range(len(data_range)):
        fila_data = data_range[i]
        val_fecha = fila_data[8] # Columna I (index 8)
        fecha = convertir_fecha(val_fecha)

        if fecha and dia_inicio <= fecha.day <= dia_fin:
            def leer(col_excel):
                v = fila_data[col_excel - 1]
                try: return int(float(v)) if v is not None else 0
                except: return 0

            registro = {
                "OBJ_FECHA": fecha,
                "PRE_CONTACTOS": leer(12), "CONTACTOS": leer(13), "PROSPECTOS": leer(14),
                "SOL_DATOS_COMPLETOS": leer(47), "VIABLES": leer(51), "CITAS_AGENDADAS": leer(15),
                "CITAS_REALES": leer(16), "DOC_COMPLETA": leer(53), "AUTORIZADAS": leer(56),
                "PEDIDO_ANTICIPO": leer(61), "DEMOS": leer(45), "ENTREGAS": leer(65), "DESEMBOLSOS": leer(63)
            }
            guardar_en_strapi(registro)
            encontrados += 1
    print(f"Bloque {dia_inicio}-{dia_fin}: {encontrados} procesados.")

def hilo_teclado(dia):
    pythoncom.CoInitialize()
    try:
        shell = win32com.client.Dispatch("WScript.Shell")
        time.sleep(2.5) 
        if shell.AppActivate("CALENDARIO") or shell.AppActivate("Microsoft Excel"):
            pestañas = {1: 0, 8: 7, 15: 14, 22: 21, 29: 28}
            saltos = pestañas.get(dia, 0)
            shell.SendKeys("{TAB 4}")
            if saltos > 0: shell.SendKeys(f"{{TAB {saltos}}}")
            shell.SendKeys("{ENTER}")
            time.sleep(1.0) 
            shell.SendKeys("%{F4}")
    finally:
        pythoncom.CoUninitialize()

def ejecutar():
    if not os.path.exists(NOMBRE_ARCHIVO):
        print("Archivo no encontrado.")
        return

    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False

    try:
        wb = excel.Workbooks.Open(NOMBRE_ARCHIVO)
        ws = wb.Sheets(HOJA_DATOS)
        puntos_inicio = [1, 8, 15, 22]
        if DIA_LIMITE >= 29: puntos_inicio.append(29)

        for inicio in puntos_inicio:
            if inicio > DIA_LIMITE: continue
            print(f"\n--- Procesando Bloque día {inicio} ---")
            
            t = threading.Thread(target=hilo_teclado, args=(inicio,))
            t.start()
            
            try: excel.Application.Run(MACRO_CALENDARIO)
            except: pass
            
            t.join()
            # ESPERA DE SEGURIDAD para que Excel termine de escribir
            time.sleep(2.0) 
            extraer_bloque_posicional(ws, inicio, DIA_LIMITE)

        wb.Save()
        print("\nProceso finalizado con éxito.")
    finally:
        try: wb.Close(False)
        except: pass
        excel.Quit()
        gc.collect()

if __name__ == "__main__":
    ejecutar()