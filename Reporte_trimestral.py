import win32com.client
import time
import os
import requests
import sys
import calendar
from datetime import datetime
import gc
import win32event
import win32api
import winerror
import pyautogui

# ================= MUTEX =================
mutex = win32event.CreateMutex(None, False, "Global\\REPORTE_TRIMESTRAL_FIX_FINAL")
if win32api.GetLastError() == winerror.ERROR_ALREADY_EXISTS:
    sys.exit(0)

# ================= PARÁMETROS =================
MES_INICIO = int(sys.argv[1]) if len(sys.argv) > 1 else 1 
ANIO = int(sys.argv[2]) if len(sys.argv) > 2 else 2026
TOKEN = sys.argv[3] if len(sys.argv) > 3 else None
SUCURSAL = (sys.argv[4] if len(sys.argv) > 4 else "Zacatecas").strip()

API_BASE = "http://localhost:1337/api"
STRAPI_URL_APVS = f"{API_BASE}/apvs"
API_URL_SUCURSALES = f"{API_BASE}/sucursals"

session = requests.Session()
if TOKEN:
    session.headers.update({"Authorization": TOKEN, "Content-Type": "application/json"})

DOC_ID_SUCURSAL = None
NOMBRE_ARCHIVO = None

# ================= CONFIG =================
def obtener_configuracion():
    global DOC_ID_SUCURSAL, NOMBRE_ARCHIVO
    try:
        res = session.get(f"{API_URL_SUCURSALES}?filters[Sucursal][$eq]={SUCURSAL}")
        if res.ok:
            data = res.json().get("data", [])
            if data:
                DOC_ID_SUCURSAL = data[0].get("documentId")
                archivo_db = data[0].get("Plantilla")
                NOMBRE_ARCHIVO = os.path.join(os.path.expanduser("~"), "Downloads", archivo_db)
                return True
    except:
        pass
    return False

# ================= FUNCIONES =================
def esperar_excel(excel):
    while excel.CalculationState != 0:
        time.sleep(0.5)

def seleccionar_mes_en_form(m):
    time.sleep(1)
    pyautogui.press('tab')
    pyautogui.press('up', presses=15)

    for _ in range(m - 1):
        pyautogui.press('down')

    pyautogui.press('tab')
    pyautogui.press('up', presses=15)

    for _ in range(ANIO - 2023):
        pyautogui.press('down')

    pyautogui.press('tab')
    pyautogui.press('enter')
    time.sleep(3)
    pyautogui.press('enter')

def extraer_bloque(ws):
    try:
        vals = ws.Range(ws.Cells(6, 186), ws.Cells(13, 186)).Value

        def f(v):
            try:
                return int(float(v)) if v else 0
            except:
                return 0

        return [f(v[0]) for v in vals]
    except:
        return None

# ================= MAIN =================
def ejecutar():

    if not obtener_configuracion() or not os.path.exists(NOMBRE_ARCHIVO):
        print("Error config o archivo")
        return

    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False

    try:
        wb = excel.Workbooks.Open(os.path.abspath(NOMBRE_ARCHIVO), UpdateLinks=0)
        ws = wb.Sheets("Reportes")

        excel.Application.Run("IrAReportes_SeguroConPwd")
        excel.Application.Run("MostrarRangoTRIMESTRAXAPV")

        meses = [MES_INICIO, MES_INICIO + 1, MES_INICIO + 2]

        for m in meses:

            print(f"\n===== MES {m} =====")

            # ---------------- CAMBIAR MES ----------------
            excel.Application.Run("MostrarUserFormAGENCIAMESAÑO")
            seleccionar_mes_en_form(m)

            esperar_excel(excel)
            wb.Save()  # 🔥 CONFIRMA CAMBIO DE MES

            # ---------------- FECHAS ----------------
            u_dia = calendar.monthrange(ANIO, m)[1]

            ws.Range("GC2").Value = f"01-{m:02d}-{ANIO}*"
            ws.Range("GC3").Value = f"{u_dia:02d}-{m:02d}-{ANIO}*"

            esperar_excel(excel)
            wb.Save()  # 🔥 CONFIRMA FECHAS

            # ---------------- MACROS MES ----------------
            excel.Application.Run("CambiarColorAzul_Boton9")
            excel.Application.Run("CambiarColorAzul_Boton11")

            esperar_excel(excel)
            wb.Save()  # 🔥 CONFIRMA RESULTADO

            data_mes = extraer_bloque(ws)
            print("MES:", data_mes)

            # ---------------- MADURACIÓN ----------------
            excel.Application.Run("CambiarColorAzul_Boton10")
            excel.Application.Run("CambiarColorAzul_Boton12")

            esperar_excel(excel)
            wb.Save()  # 🔥 CONFIRMA MADURACIÓN

            data_mad = extraer_bloque(ws)
            print("MAD:", data_mad)

            # ---------------- API ----------------
            payload = {
                "Fecha_inicio": datetime(ANIO, m, 1).strftime("%Y-%m-%d"),
                "Fecha_fin": datetime(ANIO, m, u_dia).strftime("%Y-%m-%d"),
                "tipo_registro": "TRIMESTRAL",
                "Apv_nombre": "GLOBAL_TRIMESTRE",
                "Gerente": "SISTEMAS",
                "sucursal": DOC_ID_SUCURSAL,
                "Mes": data_mes,
                "Maduracion": data_mad
            }

            try:
                session.post(STRAPI_URL_APVS, json={"data": payload})
                print(f"✔ Guardado mes {m}")
            except Exception as e:
                print("Error API:", e)

        wb.Save()

    finally:
        if 'wb' in locals():
            wb.Close(False)

        excel.Quit()
        gc.collect()

# ================= RUN =================
if __name__ == "__main__":
    ejecutar()