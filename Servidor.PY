from flask import Flask, request, jsonify
from flask_cors import CORS
import subprocess
import os

app = Flask(__name__)
CORS(app)

import subprocess
import threading

@app.route('/ejecutar-reporte', methods=['POST'])
def correr():
    
    data = request.json
    dia = data.get('dia', 31)
    sucursal = data.get('sucursal', 'Zacatecas')
    token = request.headers.get('Authorization')

    script_path = os.path.join(os.path.dirname(__file__), "reporte_nissan.py")
    print(token)

    comando = f'python "{script_path}" {dia} "{token}" "{sucursal}"'

    def ejecutar_script():
        subprocess.run(comando, shell=True)

    threading.Thread(target=ejecutar_script).start()

    return jsonify({"status": "ok"}), 200

@app.route('/ejecutar-gerente', methods=['POST'])
def correr_gerente():
    data = request.json
    dia = data.get('dia', 31)
    sucursal = data.get('sucursal', 'Zacatecas')
    token = request.headers.get('Authorization')

    # Asegúrate de que el nombre del archivo sea EXACTO al que tienes en la carpeta
    script_path = os.path.join(os.path.dirname(__file__), "reporte_global_gerente.py")
    
    # Imprimimos el comando en la consola de Flask para que lo veas y lo puedas copiar/pegar
    comando = f'python "{script_path}" {dia} "{token}" "{sucursal}"'
    print(f"DEBUG: Ejecutando comando: {comando}")

    def ejecutar_script():
        # Capturamos la salida para ver por qué no arranca
        resultado = subprocess.run(comando, shell=True, capture_output=True, text=True)
        if resultado.returncode != 0:
            print("ERROR EN EL SCRIPT:")
            print(resultado.stderr)
        else:
            print("SALIDA DEL SCRIPT:")
            print(resultado.stdout)

    threading.Thread(target=ejecutar_script).start()
    return jsonify({"status": "ok", "mensaje": "Hilo iniciado"}), 200


@app.route('/ejecutar-apv', methods=['POST'])
def correr_apv():
    data = request.json
    dia = data.get('dia', 31)
    sucursal = data.get('sucursal', 'Zacatecas')
    token = request.headers.get('Authorization')

    # Ruta al nuevo script
    script_path = os.path.join(os.path.dirname(__file__), "Reporte_apv.py")
    
    # El comando pasa Día, Token y Sucursal
    comando = f'python "{script_path}" {dia} "{token}" "{sucursal}"'

    def ejecutar_script():
        subprocess.run(comando, shell=True)

    threading.Thread(target=ejecutar_script).start()

    return jsonify({"status": "ok", "mensaje": f"Reporte APV iniciado para {sucursal}"}), 200

@app.route('/actualizar-datos', methods=['POST'])
def actualizar():
    data = request.json

    mes = data.get('mes')
    anio = data.get('anio')
    sucursal = data.get('sucursal')

    token = request.headers.get('Authorization')

    script = os.path.join(os.path.dirname(__file__), "actualizar_datos.py")

    comando = f'python "{script}" {mes} {anio} "{token}" "{sucursal}"'

    def run():
        subprocess.run(comando, shell=True)

    threading.Thread(target=run).start()

    return jsonify({"ok": True})
    
if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=False, use_reloader=False)
    
    
