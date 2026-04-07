import win32com.client
import time
import os
import requests
from datetime import datetime, timedelta

# --- CONFIGURACIÓN ---
RUTA_EXCEL = r"C:\Users\jeanl\Downloads\jgyn5k0aZuBmDgkvmusCNA8h.xlsm"
HOJA_NOMBRE = "Reportes"
STRAPI_URL = "http://localhost:1337/api/apvs"

TOKEN = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpZCI6MSwiaWF0IjoxNzc0MjM0MDc3LCJleHAiOjE3NzY4MjYwNzd9.oZS__AyaVQZoGU-dTCJpLNho-xzc7waquuWf2p5wOEw"

MACROS_MES = ["CambiarColorAzul_Boton9", "CambiarColorAzul_Boton11"]
MACROS_MAD = ["CambiarColorAzul_Boton10", "CambiarColorAzul_Boton12"]

# SESSION (🔥 mejora brutal)
session = requests.Session()
session.headers.update({
    "Authorization": f"Bearer {TOKEN}",
    "Content-Type": "application/json"
})

# CACHE (🔥 evita GET repetidos)
cache = {}

def wait_excel(excel):
    while not excel.Ready:
        time.sleep(0.05)

def formato_num(v):
    if v is None:
        return 0
    return int(v) if isinstance(v, float) and v.is_integer() else v

# EXTRAER BLOQUE
def extraer_bloque(ws, fila):
    vals = ws.Range(ws.Cells(fila,186), ws.Cells(fila+7,186)).Value
    return {
        "citas_asistidas": formato_num(vals[0][0]),
        "solicitudes_cdatos": formato_num(vals[1][0]),
        "doc_compl_mes": formato_num(vals[2][0]),
        "autorizadas": formato_num(vals[3][0]),
        "pedidos_canticipo": formato_num(vals[4][0]),
        "facturadas": formato_num(vals[5][0]),
        "desenbolsadas": formato_num(vals[6][0]),
        "entregas": formato_num(vals[7][0])
    }

# EXTRAER VENDEDORES
def extraer_vendedores(ws, fila_inicio):
    vendedores = []
    f_v = fila_inicio
    cuenta_v = 0

    while True:
        v_nom = ws.Cells(f_v, 171).Value

        if not v_nom or str(v_nom).strip() == "" or "TOTAL" in str(v_nom).upper():
            break

        nombre = str(v_nom).strip()
        data = extraer_bloque(ws, f_v + 1)

        vendedores.append({
            "nombre": nombre,
            "data": data
        })

        cuenta_v += 1

        if cuenta_v == 1:
            f_v += 10
        elif cuenta_v == 2:
            f_v += 11
        else:
            f_v += 10

    return vendedores

# VALIDACIÓN
def validar_cambios(actual, nuevo):

    hubo_cambios = False

    for tipo in ["Mes", "Maduracion"]:

        actual_data = actual.get(tipo, {}) or {}
        nuevo_data = nuevo.get(tipo, {}) or {}

        for key in nuevo_data:

            viejo = actual_data.get(key, 0)
            nuevo_val = nuevo_data.get(key, 0)

            try:
                viejo = int(viejo)
                nuevo_val = int(nuevo_val)
            except:
                pass

            if viejo > 0 and nuevo_val == 0:
                nuevo_data[key] = viejo
                continue

            if viejo != nuevo_val:
                hubo_cambios = True

        nuevo[tipo] = nuevo_data

    return hubo_cambios

# GUARDAR / ACTUALIZAR
def guardar_o_actualizar(payload):

    try:
        key = f"{payload['Fecha_inicio']}_{payload['Fecha_fin']}_{payload['tipo_registro']}_{payload['Apv_nombre']}"

        if key in cache:
            existente = cache[key]
        else:
            params = {
                "filters[Fecha_inicio][$eq]": payload["Fecha_inicio"],
                "filters[Fecha_fin][$eq]": payload["Fecha_fin"],
                "filters[tipo_registro][$eq]": payload["tipo_registro"],
                "filters[Apv_nombre][$eq]": payload["Apv_nombre"]
            }

            response = session.get(STRAPI_URL, params=params)
            existente = response.json()["data"]
            cache[key] = existente

        if existente:

            registro = existente[0]
            id_registro = registro["id"]

            if not validar_cambios(registro, payload):
                print(f"[SKIP] {payload['Apv_nombre']}")
                return

            update_url = f"{STRAPI_URL}/{id_registro}"
            session.put(update_url, json={"data": payload})

            print(f"[UPDATE] {payload['tipo_registro']} - {payload['Apv_nombre']}")

        else:
            session.post(STRAPI_URL, json={"data": payload})
            print(f"[CREATE] {payload['tipo_registro']} - {payload['Apv_nombre']}")

    except Exception as e:
        print(f"[ERROR] {e}")

# MAIN
def ejecutar():

    try:
        dia_limite = int(input("Día límite: "))
    except:
        return

    excel = None

    try:
        print("[1] Iniciando Excel...")

        excel = win32com.client.DispatchEx("Excel.Application")

        # 🔥 boost brutal
        excel.ScreenUpdating = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False

        wb = excel.Workbooks.Open(os.path.abspath(RUTA_EXCEL))
        ws = wb.Sheets(HOJA_NOMBRE)

        excel.Application.Run("MostrarRangoTRIMESTRAXAPV")
        wait_excel(excel)

        dv = ws.Range("FO15").Validation
        gerentes = [str(c.Value).strip() for c in ws.Application.Range(dv.Formula1) if c.Value]

        hoy = datetime.now()
        f_inicio_fija = hoy.replace(day=1).strftime("%Y-%m-%d")
        f_tope = hoy.replace(day=min(dia_limite, hoy.day - 1))

        for g in gerentes:

            if "POR ASIGNAR GERENTE" in g.upper():
                break

            print(f"\nGERENTE: {g}")

            ws.Cells(15, 171).Value = g
            wait_excel(excel)

            f_ini_cursor = hoy.replace(day=1)

            while f_ini_cursor <= f_tope:

                f_fin_ciclo = f_ini_cursor + timedelta(days=2)

                if f_fin_ciclo.weekday() == 6:
                    f_fin_ciclo += timedelta(days=1)

                if f_fin_ciclo > f_tope:
                    break

                f_iso_fin = f_fin_ciclo.strftime("%Y-%m-%d")
                f_excel_busqueda = f_fin_ciclo.strftime("%d-%m-%Y") + "*"

                ws.Cells(3, 185).Value = f_excel_busqueda
                wait_excel(excel)

                # MES
                for m in MACROS_MES:
                    excel.Application.Run(m)
                wait_excel(excel)

                global_mes = extraer_bloque(ws, 6)
                gerente_mes = extraer_bloque(ws, 16)
                vendedores_mes = extraer_vendedores(ws, 29)

                # MAD
                for m in MACROS_MAD:
                    excel.Application.Run(m)
                wait_excel(excel)

                global_mad = extraer_bloque(ws, 6)
                gerente_mad = extraer_bloque(ws, 16)
                vendedores_mad = extraer_vendedores(ws, 29)

                # GLOBAL
                guardar_o_actualizar({
                    "tipo_registro": "GLOBAL",
                    "Apv_nombre": "GLOBAL",
                    "Fecha_inicio": f_inicio_fija,
                    "Fecha_fin": f_iso_fin,
                    "Mes": global_mes,
                    "Maduracion": global_mad,
                    "Agencia": "Zacatecas"
                })

                # GERENTE
                guardar_o_actualizar({
                    "tipo_registro": "GERENTE",
                    "Gerente": g,
                    "Apv_nombre": g,
                    "Fecha_inicio": f_inicio_fija,
                    "Fecha_fin": f_iso_fin,
                    "Mes": gerente_mes,
                    "Maduracion": gerente_mad,
                    "Agencia": "Zacatecas"
                })

                # VENDEDORES
                for i, v in enumerate(vendedores_mes):
                    guardar_o_actualizar({
                        "tipo_registro": "VENDEDOR",
                        "Gerente": g,
                        "Apv_nombre": v["nombre"],
                        "Fecha_inicio": f_inicio_fija,
                        "Fecha_fin": f_iso_fin,
                        "Mes": v["data"],
                        "Maduracion": vendedores_mad[i]["data"],
                        "Agencia": "Zacatecas"
                    })

                f_ini_cursor = f_fin_ciclo

        wb.Save()
        print("\n[FIN] Proceso completado.")

    except Exception as e:
        print(f"[FATAL ERROR] {e}")

    finally:
        if excel:
            try:
                wb.Close(SaveChanges=True)
            except:
                pass
            excel.Quit()

if __name__ == "__main__":
    ejecutar()