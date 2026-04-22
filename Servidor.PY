from flask import Flask, request, jsonify
from flask_cors import CORS
import subprocess
import os

app = Flask(__name__)
CORS(app)

import subprocess
import threading


@app.route("/ejecutar-reporte", methods=["POST"])
def correr():
    data = request.json
    dia = data.get("dia", 31)
    sucursal = data.get("sucursal", "Zacatecas")
    token = request.headers.get("Authorization")

    script_path = os.path.join(os.path.dirname(__file__), "reporte_nissan.py")
    
    comando = f'python "{script_path}" {dia} "{token}" "{sucursal}"'
    print(f"Ejecutando: {comando}")

    try:
        resultado = subprocess.run(comando, shell=True, capture_output=True, text=True)

        if resultado.returncode != 0:
            print("Error en script:", resultado.stderr)
            return jsonify({"status": "error", "error": resultado.stderr}), 500

        return jsonify({
            "status": "ok", 
            "message": "Extracción finalizada y datos guardados en Strapi"
        }), 200

    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500

@app.route("/ejecutar-gerente", methods=["POST"])
def correr_gerente():
    data = request.json
    dia = data.get("dia", 31)
    sucursal = data.get("sucursal", "Zacatecas")
    token = request.headers.get("Authorization")

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


@app.route("/ejecutar-apv", methods=["POST"])
def correr_apv():
    data = request.json
    dia = data.get("dia", 31)
    sucursal = data.get("sucursal", "Zacatecas")
    token = request.headers.get("Authorization")

    # Ruta al script de reporte APV
    script_path = os.path.join(os.path.dirname(__file__), "Reporte_apv.py")

    # Construcción del comando con argumentos escapados
    comando = f'python "{script_path}" {dia} "{token}" "{sucursal}"'
    print(f"DEBUG: Ejecutando reporte APV: {comando}")

    try:
        # Se elimina el Threading para que sea síncrono.
        # Flask esperará aquí hasta que el script termine.
        resultado = subprocess.run(
            comando, shell=True, capture_output=True, text=True, check=True
        )

        # Si llega aquí, el proceso terminó con éxito (exit code 0)
        print("SALIDA DEL SCRIPT APV:")
        print(resultado.stdout)

        return (
            jsonify(
                {
                    "ok": True,
                    "mensaje": f"Reporte APV para {sucursal} finalizado y guardado correctamente",
                }
            ),
            200,
        )

    except subprocess.CalledProcessError as e:
        # Si el script falla (por ejemplo, error de Excel o de red), entra aquí
        print("ERROR EN EL SCRIPT APV:")
        print(e.stderr)
        return (
            jsonify(
                {
                    "ok": False,
                    "error": "El proceso de reporte APV falló internamente",
                    "detalles": e.stderr,
                }
            ),
            500,
        )


@app.route("/actualizar-datos", methods=["POST"])
def actualizar():
    data = request.json
    mes = data.get("mes")
    anio = data.get("anio")
    sucursal = data.get("sucursal")
    token = request.headers.get("Authorization")

    script = os.path.join(os.path.dirname(__file__), "actualizar_datos.py")
    comando = f'python "{script}" {mes} {anio} "{token}" "{sucursal}"'

    try:
        # Ejecucion sincrona: Flask espera la finalizacion del proceso
        resultado = subprocess.run(
            comando, shell=True, capture_output=True, text=True, check=True
        )
        return jsonify({"ok": True, "mensaje": "Proceso finalizado y guardado"}), 200
    except subprocess.CalledProcessError as e:
        print(f"Error: {e.stderr}")
        return jsonify({"ok": False, "error": "Error en la ejecucion del script"}), 500


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=False, use_reloader=False)
