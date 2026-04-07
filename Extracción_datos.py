from flask import Flask, jsonify
from flask_cors import CORS
import pandas as pd
import re

app = Flask(__name__)
CORS(app) # IMPORTANTE: Permite que Ionic acceda a la API

def extraer_nissan_posicional(ruta_archivo):
    try:
        df = pd.read_excel(ruta_archivo, sheet_name='Reportes', engine='openpyxl')
        df = df.fillna(0)
        
        # 1. Usaremos una lista en lugar de un diccionario
        lista_reportes = [] 

        for index, fila in df.iterrows():
            valor_i = str(fila.iloc[8]).strip()
            valor_j = str(fila.iloc[9]).strip()
            valor_fecha = None
            if re.search(r'\d{1,2}[-/]\d{1,2}', valor_i): valor_fecha = valor_i
            elif re.search(r'\d{1,2}[-/]\d{1,2}', valor_j): valor_fecha = valor_j
            
            if valor_fecha:
                fecha_id = valor_fecha.replace('*', '').split(' ')[0]
                def val(p):
                    try: return int(float(fila.iloc[p]))
                    except: return 0

                # 2. Agregamos el campo "FECHA" dentro del objeto
                registro = {
                    "FECHA": fecha_id,  # <--- Ahora la fecha va dentro del array
                    "PRE_CONTACTOS": val(11),
                    "CONTACTOS": val(12),
                    "PROSPECTOS": val(13),
                    "SOL_DATOS_COMPLETOS": val(46),
                    "VIABLES": val(50),
                    "CITAS_AGENDADAS": val(14), 
                    "CITAS_REALES": val(15),
                    "DOC_COMPLETA": val(52),
                    "AUTORIZADAS": val(54),
                    "PEDIDO_ANTICIPO": val(60),
                    "DEMOS": val(44),
                    "ENTREGAS": val(64),     
                    "DESEMBOLSOS": val(62)
                }
                lista_reportes.append(registro)

        return lista_reportes # Retorna la lista de objetos
    except Exception as e:
        return {"error": str(e)}
# --- RUTA DE LA API ---
@app.route('/api/reporte', methods=['GET'])
def get_reporte():
    ruta = r"C:\Users\jeanl\Downloads\wzhSc3ey_LmfyiYXsJH44P_K.xlsm"
    datos = extraer_nissan_posicional(ruta)
    return jsonify(datos) # Enviamos el JSON a Ionic

if __name__ == '__main__':
    # host='0.0.0.0' permite que otros dispositivos (celular) vean la API
    app.run(debug=True, host='0.0.0.0', port=5000)