import pandas as pd
import re
import json

def extraer_nissan_posicional(ruta_archivo):
    try:
        # Carga del archivo Excel
        df = pd.read_excel(ruta_archivo, sheet_name='Reportes', engine='openpyxl')
        df = df.fillna(0)
        
        lista_reportes = [] 

        for index, fila in df.iterrows():
            # Extraer valores de las columnas I (8) y J (9) para la fecha
            valor_i = str(fila.iloc[8]).strip()
            valor_j = str(fila.iloc[9]).strip()
            valor_fecha = None
            
            if re.search(r'\d{1,2}[-/]\d{1,2}', valor_i): 
                valor_fecha = valor_i
            elif re.search(r'\d{1,2}[-/]\d{1,2}', valor_j): 
                valor_fecha = valor_j
            
            if valor_fecha:
                # Limpieza de la fecha
                fecha_id = valor_fecha.replace('*', '').split(' ')[0]
                
                # Función interna para validar y convertir a entero
                def val(p):
                    try: 
                        return int(float(fila.iloc[p]))
                    except: 
                        return 0

                # Estructura de los datos según tus requerimientos
                registro = {
                    "fecha": fecha_id,
                    "preContactos": val(11),
                    "Contactos": val(12),
                    "prospectos": val(13),
                    "solCDatosCompletos": val(46),
                    "viablesPreAutorizadas": val(50),
                    "citasAgendadas": val(14), 
                    "citasReales": val(15),
                    "docCompleta": val(52),
                    "autorizadas": val(54),
                    "pedidosConAnticipo": val(60),
                    "demos": val(44),
                    "entregas": val(64),     
                    "desembolsos": val(62)
                }
                lista_reportes.append(registro)

        return lista_reportes 
    except Exception as e:
        return {"error": str(e)}

# --- EJECUCIÓN DIRECTA ---
if __name__ == '__main__':
    # Ruta del archivo
    ruta = r"C:\Users\jeanl\Downloads\jgyn5k0aZuBmDgkvmusCNA8h.xlsm"
    
    print(f"Extrayendo datos de: {ruta}...\n")
    datos = extraer_nissan_posicional(ruta)
    
    # Si hay un error en la extracción, lo imprime
    if isinstance(datos, dict) and "error" in datos:
        print(f"Error: {datos['error']}")
    else:
        # Imprime los datos con formato JSON legible
        print(json.dumps(datos, indent=4, ensure_ascii=False))
        print(f"\nTotal de registros extraídos: {len(datos)}")