import win32com.client
import pythoncom
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

# --- CONFIGURACIÓN Y MUTEX ---
mutex = win32event.CreateMutex(None, False, "Global\\REPORTE_APVS_FINAL_MUTEX")
if win32api.GetLastError() == winerror.ERROR_ALREADY_EXISTS:
    sys.exit(0)

DIA_LIMITE_SOLICITADO = int(sys.argv[1]) if len(sys.argv) > 1 else 31
TOKEN = sys.argv[2] if len(sys.argv) > 2 else None
SUCURSAL_BUSQUEDA = (sys.argv[3] if len(sys.argv) > 3 else "Zacatecas").strip()

HOJA_NOMBRE = "Reportes" 
API_BASE = "http://localhost:1337/api"
STRAPI_URL_APVS = f"{API_BASE}/apvs"
API_URL_SUCURSALES = f"{API_BASE}/sucursals"

session = requests.Session()
if TOKEN:
    session.headers.update({"Authorization": f"{TOKEN}", "Content-Type": "application/json"})

DOC_ID_SUCURSAL = None
NOMBRE_ARCHIVO = None

def obtener_configuracion_sucursal():
    global DOC_ID_SUCURSAL, NOMBRE_ARCHIVO
    try:
        res = session.get(f"{API_URL_SUCURSALES}?filters[Sucursal][$eq]={SUCURSAL_BUSQUEDA}")
        if res.ok:
            datos = res.json().get('data', [])
            if datos:
                suc = datos[0]
                DOC_ID_SUCURSAL = suc.get('documentId')
                archivo_db = suc.get('Plantilla') 
                path_dl = os.path.join(os.path.expanduser("~"), "Downloads")
                NOMBRE_ARCHIVO = os.path.join(path_dl, archivo_db)
                return True
        return False
    except Exception: return False

def calcular_dias_corte(anio, mes, limite):
    dias = []
    u_dia_mes = calendar.monthrange(anio, mes)[1]
    actual = 3
    while actual <= limite:
        f_temp = datetime(anio, mes, actual)
        if f_temp.weekday() == 6: actual += 1
        if actual > limite: break
        dias.append(actual)
        actual += 2
    if limite >= (u_dia_mes - 1): 
        if not dias or dias[-1] != limite:
            dias.append(limite)
    return dias

def extraer_bloque(ws, fila):
    vals = ws.Range(ws.Cells(fila, 186), ws.Cells(fila + 7, 186)).Value
    def f_n(v):
        try: return int(float(v)) if v is not None else 0
        except: return 0
    return {
        "citas_asistidas": f_n(vals[0][0]), "solicitudes_cdatos": f_n(vals[1][0]),
        "doc_compl_mes": f_n(vals[2][0]), "autorizadas": f_n(vals[3][0]),
        "pedidos_canticipo": f_n(vals[4][0]), "facturadas": f_n(vals[5][0]),
        "desenbolsadas": f_n(vals[6][0]), "entregas": f_n(vals[7][0])
    }

def guardar_o_actualizar(payload):
    try:
        params = {
            "filters[Fecha_fin][$eq]": payload["Fecha_fin"],
            "filters[tipo_registro][$eq]": payload["tipo_registro"],
            "filters[Apv_nombre][$eq]": payload["Apv_nombre"],
            "filters[sucursal][documentId][$eq]": DOC_ID_SUCURSAL
        }
        res_get = session.get(STRAPI_URL_APVS, params=params)
        data = res_get.json().get("data", [])
        if data:
            session.put(f"{STRAPI_URL_APVS}/{data[0]['documentId']}", json={"data": payload})
        else:
            session.post(STRAPI_URL_APVS, json={"data": payload})
    except Exception: pass

def ejecutar():
    if not obtener_configuracion_sucursal() or not os.path.exists(NOMBRE_ARCHIVO): return

    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.EnableEvents = False # CAMBIO: Debe estar en True para que las macros de los botones reaccionen
    excel.AutomationSecurity = 1 

    try:
        wb = excel.Workbooks.Open(os.path.abspath(NOMBRE_ARCHIVO), UpdateLinks=0)
        ws = wb.Sheets(HOJA_NOMBRE)

        # 1. Preparar entorno
        try:
            excel.Application.Run("IrAReportes_SeguroConPwd")
            excel.Application.Run("MostrarRangoTRIMESTRAXAPV")
        except: pass

        # 2. Configurar Fechas Base
        celda_inicio = ws.Range("GC2")
        # ... (lógica de validación de fecha que ya tienes)
        
        raw_f = str(celda_inicio.Value).replace("*", "").strip()
        f_base = datetime.strptime(raw_f, "%d-%m-%Y")
        f_ini_iso = f_base.replace(day=1).strftime("%Y-%m-%d")
        dias_corte = calcular_dias_corte(f_base.year, f_base.month, DIA_LIMITE_SOLICITADO)

        # 3. Ciclo de Gerentes
        form_g = ws.Range("FO15").Validation.Formula1
        gerentes = [str(c.Value).strip() for c in ws.Application.Range(form_g) if c.Value]

        for g in gerentes:
            if "POR ASIGNAR" in g.upper() or not g: continue
            ws.Cells(15, 171).Value = g # Cambia Gerente
            
            for d in dias_corte:
                f_fin_iso = datetime(f_base.year, f_base.month, d).strftime("%Y-%m-%d")
                ws.Range("GC3").Value = datetime(f_base.year, f_base.month, d).strftime("%d-%m-%Y") + "*"
                
                # --- PASO A: EXTRAER "MES" ---
                try:
                    excel.Application.Run("CambiarColorAzul_Boton9")  # Global Mes
                    excel.Application.Run("CambiarColorAzul_Boton11") # APV Mes
                    excel.Calculate() # Forzar a Excel a actualizar celdas
                    time.sleep(1)     # Pausa técnica para estabilidad
                except: pass

                data_mes_global = extraer_bloque(ws, 6)
                data_mes_gerente = extraer_bloque(ws, 16)
                
                # Vendedores (Mes)
                v_list_mes = []
                f_v, cv = 29, 0
                while True:
                    nom = ws.Cells(f_v, 171).Value
                    if not nom or "TOTAL" in str(nom).upper(): break
                    v_list_mes.append({"n": str(nom).strip(), "d": extraer_bloque(ws, f_v + 1)})
                    cv += 1
                    f_v += 11 if cv == 2 else 10

                # --- PASO B: EXTRAER "MADURACIÓN" ---
                try:
                    excel.Application.Run("CambiarColorAzul_Boton10") # Global Maduracion
                    excel.Application.Run("CambiarColorAzul_Boton12") # APV Maduracion
                    excel.Calculate() # Forzar a Excel a actualizar celdas
                    time.sleep(1)
                except: pass

                data_mad_global = extraer_bloque(ws, 6)
                data_mad_gerente = extraer_bloque(ws, 16)
                
                # Vendedores (Maduración)
                v_dict_mad = {}
                f_v, cv = 29, 0
                while True:
                    nom = ws.Cells(f_v, 171).Value
                    if not nom or "TOTAL" in str(nom).upper(): break
                    v_dict_mad[str(nom).strip()] = extraer_bloque(ws, f_v + 1)
                    cv += 1
                    f_v += 11 if cv == 2 else 10

                base = {"Fecha_inicio": f_ini_iso, "Fecha_fin": f_fin_iso, "sucursal": DOC_ID_SUCURSAL}
                
                # Global
                guardar_o_actualizar({**base, "tipo_registro": "GLOBAL", "Apv_nombre": "GLOBAL", "Gerente": "GLOBAL", 
                                      "Mes": data_mes_global, "Maduracion": data_mad_global})
                # Gerente
                guardar_o_actualizar({**base, "tipo_registro": "GERENTE", "Gerente": g, "Apv_nombre": g, 
                                      "Mes": data_mes_gerente, "Maduracion": data_mad_gerente})
                # Vendedores
                for v in v_list_mes:
                    nombre = v["n"]
                    guardar_o_actualizar({**base, "tipo_registro": "VENDEDOR", "Gerente": g, "Apv_nombre": nombre, 
                                          "Mes": v["d"], "Maduracion": v_dict_mad.get(nombre, v["d"])})

        print("Sincronización exitosa con distinción de Mes/Maduración.")
    finally:
        if 'wb' in locals(): wb.Close(False)
        excel.Quit()
        gc.collect()

if __name__ == "__main__":
    ejecutar()