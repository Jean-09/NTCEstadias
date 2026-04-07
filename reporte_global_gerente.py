import win32com.client
import pythoncom
import time
import os
import re
import threading
import sys
from datetime import datetime
import gc
import win32event
import win32api
import winerror

# --- PREVENIR DOBLE INSTANCIA ---
mutex = win32event.CreateMutex(None, False, "Global\\REPORTE_NISSAN_STEP_BY_STEP")
if win32api.GetLastError() == winerror.ERROR_ALREADY_EXISTS:
    print("Ya hay una instancia ejecutándose. Cancelado.")
    sys.exit(0)

# --- CONFIGURACIÓN ---
DIA_LIMITE = int(sys.argv[1]) if len(sys.argv) > 1 else 31
SUCURSAL = (sys.argv[3] if len(sys.argv) > 3 else "Zacatecas").strip()

if SUCURSAL.lower() == "bicentenario":
    NOMBRE_ARCHIVO = r"C:\Users\jeanl\Downloads\jgyn5k0aZuBmDgkvmusCNA8h.xlsm"
else:
    NOMBRE_ARCHIVO = r"C:\Users\jeanl\Downloads\4mYaHHSktW5igUO5zi6FEYU6.xlsm"

HOJA_DATOS = "Reportes"
CELDA_NOMBRE_GERENTE = "B2"  # Celda donde se refleja el nombre seleccionado

# --- FUNCIONES DE APOYO ---

def convertir_fecha(valor):
    if not valor: return None
    texto = str(valor).strip()
    match = re.search(r'(\d{1,2})[-/](\d{1,2})', texto)
    if match:
        dia, mes = int(match.group(1)), int(match.group(2))
        if mes == datetime.now().month:
            return datetime(datetime.now().year, mes, dia)
    return None

def extraer_datos_gerente_dia(ws, dia_inicio, dia_limite, nombre_g):
    ultima_fila = ws.UsedRange.Rows.Count
    dia_fin = dia_limite if dia_inicio == 22 else min(dia_inicio + 6, dia_limite)
    mes_act = datetime.now().month
    encontrados = 0

    for fila in range(2, ultima_fila + 1):
        v9 = ws.Cells(fila, 9).Value
        v10 = ws.Cells(fila, 10).Value
        fecha = convertir_fecha(v9) or convertir_fecha(v10)

        if fecha and (dia_inicio <= fecha.day <= dia_fin) and fecha.month == mes_act:
            def val(p):
                try:
                    v = ws.Cells(fila, p).Value
                    return int(float(v)) if v else 0
                except: return 0

            f_str = fecha.strftime('%Y-%m-%d')
            citas = val(16)
            entregas = val(65)
            print(f"      [DATOS] Gerente: {nombre_g} | Fecha: {f_str} | Citas: {citas} | Entregas: {entregas}")
            encontrados += 1
    
    if encontrados == 0:
        print(f"      [!] No se encontraron datos en el rango del día {dia_inicio} al {dia_fin}")

def hilo_control_calendario(dia):
    """Maneja la interacción con el modal de calendario (MostrarCalendario)"""
    pythoncom.CoInitialize()
    try:
        shell = win32com.client.Dispatch("WScript.Shell")
        time.sleep(2.5) # Espera a que el UserForm de calendario aparezca
        
        if shell.AppActivate("CALENDARIO") or shell.AppActivate("Microsoft Excel"):
            time.sleep(0.5)
            shell.SendKeys("{TAB 4}") # Navegar a los botones
            
            if dia == 8: shell.SendKeys("{TAB 7}")
            elif dia == 15: shell.SendKeys("{TAB 14}")
            elif dia == 22: shell.SendKeys("{TAB 21}")
            elif dia == 29: shell.SendKeys("{TAB 28}")
            
            time.sleep(0.5)
            shell.SendKeys("{ENTER}") # Ejecutar selección de semana
            time.sleep(1.0)
            shell.SendKeys("%{F4}") # Cerrar el modal del calendario
    finally:
        pythoncom.CoUninitialize()

def ejecutar():
    if not os.path.exists(NOMBRE_ARCHIVO):
        print(f"ERROR: Archivo no encontrado en {NOMBRE_ARCHIVO}")
        return

    excel = None
    try:
        # 1. Iniciar Excel
        print("1. Iniciando Excel...")
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = True
        excel.DisplayAlerts = False
        
        wb = excel.Workbooks.Open(NOMBRE_ARCHIVO)
        ws = wb.Sheets(HOJA_DATOS)
        shell = win32com.client.Dispatch("WScript.Shell")

        # 2. Ubicar Rango Calendario
        print("2. Ejecutando MostrarRangoCALENDARIO...")
        excel.Application.Run("MostrarRangoCALENDARIO")
        time.sleep(2)
        print("   -> Se abrio el rango de calendario.")

        # 3. Abrir Modal Gerentes
        print("3. Ejecutando CambiarColorAzul_Botono3 (Modal Gerentes)...")
        excel.Application.Run("CambiarColorAzul_Botono3")
        time.sleep(3)
        print("   -> Modal de gerentes abierto.")
        
        gerentes_vistos = []

        # 4. Bucle Principal de Gerentes
        for i in range(20): # Límite de gerentes
            time.sleep(1.5)
            # Traer al frente el modal de gerentes para asegurar lectura y navegación
            shell.AppActivate("SELECCIONA UN GERENTE")
            
            nombre_actual = str(ws.Range(CELDA_NOMBRE_GERENTE).Value).strip()
            
            if (gerentes_vistos and nombre_actual == gerentes_vistos[-1]) or nombre_actual in ["None", ""]:
                print("\n*** Fin de la lista de gerentes o celda vacia. ***")
                break
            
            gerentes_vistos.append(nombre_actual)
            print(f"\n>>> TRABAJANDO CON GERENTE: {nombre_actual}")

            # 5. Por cada Gerente, recorrer sus semanas
            bloques = [b for b in [1, 8, 15, 22] if b <= DIA_LIMITE]
            if DIA_LIMITE > 28: bloques.append(29)

            for inicio in bloques:
                print(f"   -> Abriendo MostrarCalendario para bloque dia {inicio}...")
                
                # Lanzamos el hilo que cerrará el calendario DESPUÉS de que la macro lo abra
                t = threading.Thread(target=hilo_control_calendario, args=(inicio,))
                t.start()
                
                try:
                    # Esta llamada es bloqueante hasta que el modal se cierra, 
                    # por eso el hilo de arriba es el que lo gestiona por fuera.
                    excel.Application.Run("MostrarCalendario")
                except Exception as e:
                    pass
                
                t.join() # Esperamos a que el hilo termine de interactuar
                time.sleep(1.5)
                
                # Extraer y mostrar datos en consola
                extraer_datos_gerente_dia(ws, inicio, DIA_LIMITE, nombre_actual)

            # 6. Al terminar el gerente actual, bajar al siguiente en el modal
            print(f"   -> Finalizado {nombre_actual}. Bajando al siguiente gerente...")
            shell.AppActivate("SELECCIONA UN GERENTE")
            time.sleep(0.5)
            shell.SendKeys("{DOWN}")

        # 7. Cerrar modal de gerentes
        print("\n4. Cerrando modal de gerentes y finalizando.")
        shell.AppActivate("SELECCIONA UN GERENTE")
        shell.SendKeys("%{F4}")

    except Exception as e:
        print(f"ERROR DURANTE LA EJECUCIÓN: {e}")
    finally:
        if excel:
            # wb.Save() # Opcional
            gc.collect()

if __name__ == "__main__":
    ejecutar()