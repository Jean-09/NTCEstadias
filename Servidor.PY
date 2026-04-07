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
    
if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=False, use_reloader=False)