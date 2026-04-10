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

    try:
        wb = excel.Workbooks.Open(os.path.abspath(NOMBRE_ARCHIVO))
        ws = wb.Sheets(HOJA_NOMBRE)
        raw_f = str(ws.Cells(2, 185).Value).replace("*", "").strip()
        f_base = datetime.strptime(raw_f, "%d-%m-%Y")
        f_ini_iso = f_base.replace(day=1).strftime("%Y-%m-%d")
        dias_corte = calcular_dias_corte(f_base.year, f_base.month, DIA_LIMITE_SOLICITADO)

        gerentes = [str(c.Value).strip() for c in ws.Application.Range(ws.Range("FO15").Validation.Formula1) if c.Value]

        for g in gerentes:
            if "POR ASIGNAR" in g.upper() or not g: continue
            ws.Cells(15, 171).Value = g

            for d in dias_corte:
                f_fin_iso = datetime(f_base.year, f_base.month, d).strftime("%Y-%m-%d")
                ws.Cells(3, 185).Value = datetime(f_base.year, f_base.month, d).strftime("%d-%m-%Y") + "*"
                
                # --- EXTRACCIÓN DATOS MES ---
                try:
                    for m in ["Boton9", "Boton11"]: excel.Application.Run(f"CambiarColorAzul_{m}")
                except: pass
                
                mes_global = extraer_bloque(ws, 6)
                mes_gerente = extraer_bloque(ws, 16)
                v_mes_list = []
                f_v, cv = 29, 0
                while True:
                    nom = ws.Cells(f_v, 171).Value
                    if not nom or "TOTAL" in str(nom).upper(): break
                    v_mes_list.append({"n": str(nom).strip(), "d": extraer_bloque(ws, f_v + 1)})
                    cv += 1
                    f_v += 11 if cv == 2 else 10

                # --- EXTRACCIÓN DATOS MADURACIÓN ---
                try:
                    for m in ["Boton10", "Boton12"]: excel.Application.Run(f"CambiarColorAzul_{m}")
                except: pass

                mad_global = extraer_bloque(ws, 6)
                mad_gerente = extraer_bloque(ws, 16)
                v_mad_dict = {}
                f_v, cv = 29, 0
                while True:
                    nom = ws.Cells(f_v, 171).Value
                    if not nom or "TOTAL" in str(nom).upper(): break
                    v_mad_dict[str(nom).strip()] = extraer_bloque(ws, f_v + 1)
                    cv += 1
                    f_v += 11 if cv == 2 else 10

                # --- GUARDADO FINAL ---
                base = {"Fecha_inicio": f_ini_iso, "Fecha_fin": f_fin_iso, "sucursal": DOC_ID_SUCURSAL}
                
                guardar_o_actualizar({**base, "tipo_registro": "GLOBAL", "Apv_nombre": "GLOBAL", "Gerente": "GLOBAL", "Mes": mes_global, "Maduracion": mad_global})
                guardar_o_actualizar({**base, "tipo_registro": "GERENTE", "Gerente": g, "Apv_nombre": g, "Mes": mes_gerente, "Maduracion": mad_gerente})
                
                for v in v_mes_list:
                    nom_v = v["n"]
                    guardar_o_actualizar({**base, "tipo_registro": "VENDEDOR", "Gerente": g, "Apv_nombre": nom_v, "Mes": v["d"], "Maduracion": v_mad_dict.get(nom_v, v["d"])})

        wb.Save()
        print("✅ Sincronización exitosa.")
    finally:
        if 'wb' in locals(): wb.Close(False)
        excel.Quit()
        gc.collect()

if __name__ == "__main__":
    ejecutar()