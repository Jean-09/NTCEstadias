from flask import Flask, request, jsonify
from flask_cors import CORS
import subprocess
import os

app = Flask(__name__)
CORS(app)

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
            return jsonify({"status": "error", "error": resultado.stderr}), 500
        return jsonify({"status": "ok", "message": "Finalizado"}), 200
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500

@app.route("/ejecutar-gerente", methods=["POST"])
def correr_gerente():
    data = request.json
    dia = data.get("dia", 31)
    sucursal = data.get("sucursal", "Zacatecas")
    token = request.headers.get("Authorization")

    script_path = os.path.join(os.path.dirname(__file__), "reporte_global_gerente.py")
    comando = f'python "{script_path}" {dia} "{token}" "{sucursal}"'
    
    print(f"DEBUG: Iniciando extracción pesada de Gerentes para {sucursal}...")

    try:
        # Quitamos el Threading. Ahora Flask esperará aquí "eternamente" hasta que termine.
        resultado = subprocess.run(
            comando, 
            shell=True, 
            capture_output=True, 
            text=True, 
            check=True # Esto lanza una excepción si el código de salida no es 0
        )

        print("SALIDA GERENTE:")
        print(resultado.stdout)

        return jsonify({
            "ok": True, 
            "mensaje": f"Reporte de Gerentes para {sucursal} finalizado exitosamente"
        }), 200

    except subprocess.CalledProcessError as e:
        print("ERROR EN EL SCRIPT DE GERENTE:")
        print(e.stderr)
        return jsonify({
            "ok": False, 
            "error": "Error interno en el script de Gerente",
            "detalles": e.stderr
        }), 500
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500

@app.route("/ejecutar-apv", methods=["POST"])
def correr_apv():
    data = request.json
    dia = data.get("dia", 31)
    sucursal = data.get("sucursal", "Zacatecas")
    token = request.headers.get("Authorization")

    script_path = os.path.join(os.path.dirname(__file__), "Reporte_apv.py")
    comando = f'python "{script_path}" {dia} "{token}" "{sucursal}"'
    
    print(f"DEBUG: Iniciando reporte APV para {sucursal}. Esperando finalización...")

    try:
        # Aquí Flask se detiene y espera a que el proceso termine por completo
        resultado = subprocess.run(
            comando, 
            shell=True, 
            capture_output=True, 
            text=True, 
            check=True  # Si el script truena, salta al except automáticamente
        )

        # Si llega aquí, es que terminó con éxito
        print("SALIDA EXITOSA APV:")
        print(resultado.stdout)

        return jsonify({
            "ok": True,
            "mensaje": f"Reporte APV para {sucursal} finalizado y guardado correctamente",
        }), 200

    except subprocess.CalledProcessError as e:
        # Aquí capturamos el error si el script de Python falla internamente
        print("ERROR EN EL SCRIPT APV:")
        print(e.stderr)
        return jsonify({
            "ok": False,
            "error": "El proceso de reporte APV falló internamente",
            "detalles": e.stderr,
        }), 500

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
        subprocess.run(comando, shell=True, capture_output=True, text=True, check=True)
        return jsonify({"ok": True, "mensaje": "Proceso finalizado"}), 200
    except subprocess.CalledProcessError as e:
        return jsonify({"ok": False, "error": "Error en la ejecución"}), 500

if __name__ == "__main__":
    # Importante: No uses debug=True en producción porque puede reiniciar el proceso
    app.run(host="0.0.0.0", port=5000, debug=False, use_reloader=False)