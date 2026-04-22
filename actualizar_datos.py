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

xlDone = 0

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

def obtener_archivo():
    # Gestiona la localizacion de la plantilla desde la API
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

def esperar_excel(pid_excel):
    # Monitorea el foco de la ventana de Excel
    while True:
        hwnd = win32gui.GetForegroundWindow()
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        if pid == pid_excel:
            return
        time.sleep(0.2)

def cerrar_popup():
    # Acciona el boton Aceptar de la alerta de finalizacion
    pyautogui.press('enter')
    time.sleep(0.5)

def seleccionar_mes_anio():
    # Interactua con el formulario para disparar la sincronizacion
    pyautogui.press('tab')
    pyautogui.press('up', presses=15)
    for _ in range(MES - 1):
        pyautogui.press('down')

    pyautogui.press('tab')
    pyautogui.press('up', presses=15)
    for _ in range(ANIO - ANIO_BASE):
        pyautogui.press('down')

    pyautogui.press('tab')
    pyautogui.press('enter')

def esperar_calculos(excel):
    # Verifica que el motor de calculo haya terminado sus tareas
    while True:
        try:
            if excel.CalculationState == xlDone and excel.Ready:
                break
        except:
            pass
        time.sleep(0.1)

def ejecutar():
    # Controla el flujo completo de apertura, actualizacion y guardado
    if not obtener_archivo():
        sys.exit(1)

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
    esperar_calculos(excel)
    
    cerrar_popup()
    esperar_calculos(excel)

    wb.Save()
    
    gc.collect()

if __name__ == "__main__":
    ejecutar()