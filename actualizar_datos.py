import win32com.client
import pythoncom
import time
import os
import sys
import requests
import win32event
import win32api
import winerror
import pyautogui
import win32gui
import win32process
import gc

# ================= MUTEX =================
mutex = win32event.CreateMutex(None, False, "Global\\ACTUALIZAR_DATOS_MUTEX")
if win32api.GetLastError() == winerror.ERROR_ALREADY_EXISTS:
    sys.exit(0)

# ================= PARAMETROS =================
MES = int(sys.argv[1])
ANIO = int(sys.argv[2])
TOKEN = sys.argv[3]
SUCURSAL = sys.argv[4]

ANIO_BASE = 2023

API_BASE = "http://localhost:1337/api"
API_URL_SUCURSALES = f"{API_BASE}/sucursals"

session = requests.Session()
session.headers.update({"Authorization": TOKEN})

NOMBRE_ARCHIVO = None

# ================= STRAPI =================
def obtener_archivo():
    global NOMBRE_ARCHIVO
    try:
        res = session.get(f"{API_URL_SUCURSALES}?filters[Sucursal][$eq]={SUCURSAL}")
        if res.ok:
            data = res.json().get("data", [])
            if data:
                archivo_db = data[0].get("Plantilla")
                ruta = os.path.join(os.path.expanduser("~"), "Downloads", archivo_db)
                NOMBRE_ARCHIVO = ruta
                return True
    except:
        pass
    return False

# ================= CONTROL EXCEL =================
def esperar_excel(pid_excel):
    while True:
        hwnd = win32gui.GetForegroundWindow()
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        if pid == pid_excel:
            return
        time.sleep(0.2)

def esperar_listo(excel):
    while True:
        try:
            if excel.Ready:
                break
        except:
            pass
        time.sleep(0.3)

def cerrar_popup():
    for _ in range(5):
        pyautogui.press('enter')
        time.sleep(0.2)

def seleccionar_mes_anio():
    # MES
    pyautogui.press('tab')
    pyautogui.press('up', presses=15)
    for _ in range(MES - 1):
        pyautogui.press('down')

    # AÑO
    pyautogui.press('tab')
    pyautogui.press('up', presses=15)
    for _ in range(ANIO - ANIO_BASE):
        pyautogui.press('down')

    # SINCRONIZAR
    pyautogui.press('tab')
    pyautogui.press('enter')

# ================= MAIN =================
def ejecutar():
    if not obtener_archivo():
        print("❌ No se encontró plantilla")
        return

    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False

    excel.EnableEvents = False
    wb = excel.Workbooks.Open(NOMBRE_ARCHIVO)
    excel.EnableEvents = True

    hwnd = excel.Hwnd
    _, pid_excel = win32process.GetWindowThreadProcessId(hwnd)

    excel.Application.Run("MostrarUserFormAGENCIAMESAÑO")

    esperar_excel(pid_excel)

    seleccionar_mes_anio()

    esperar_listo(excel)
    cerrar_popup()

    wb.Save()

    wb.Close(False)
    excel.Quit()
    gc.collect()

if __name__ == "__main__":
    ejecutar()