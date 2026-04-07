import os
import json
import shutil
from datetime import datetime
from src.egresos.rutas.config import ConfigRutas


def mover_si_todo_procesado_egresos():
    try:
        carpeta_proceso = ConfigRutas.FOLDER_PROCESO
        carpeta_destino_excel = ConfigRutas.FOLDER_FINAL
        carpeta_destino_json = os.path.join(os.path.dirname(carpeta_proceso), "Auditoria")

        # Leer JSON
        path_json = os.path.join(carpeta_proceso, "egresos.json")
        if not os.path.exists(path_json):
            print("No existe egresos.json, no hay nada que verificar")
            return False

        with open(path_json, "r", encoding="utf-8") as f:
            data = json.load(f)

        total = len(data)
        procesados = sum(1 for d in data if d.get("procesado") == "si")
        con_observacion = [d for d in data if d.get("procesado") == "si" and d.get("observacion")]
        pendientes = total - procesados

        print(f"Verificacion final: {procesados}/{total} procesados | {pendientes} pendientes | {len(con_observacion)} con observacion")

        if pendientes > 0:
            print(f"Aun faltan {pendientes} lotes por procesar")
            return False

        # Todos procesados pero algunos con observacion -> reprocesar (max 3 intentos)
        MAX_INTENTOS = 5
        reprocesar = []
        agotados = []

        for d in con_observacion:
            intento_actual = d.get("intento", 1)
            if intento_actual >= MAX_INTENTOS:
                agotados.append(d)
            else:
                reprocesar.append(d)

        if agotados:
            print(f"{len(agotados)} lote(s) agotaron los {MAX_INTENTOS} intentos. No se reprocesaran.")
            for d in agotados:
                print(f"  - {d.get('haber', {}).get('fecha')} | observacion: {d.get('observacion')} | intentos: {d.get('intento', 1)}")

        if reprocesar:
            print(f"{len(reprocesar)} lote(s) con observacion. Se marcaran para reprocesar.")
            for d in reprocesar:
                intento_actual = d.get("intento", 1)
                d["intento"] = intento_actual + 1
                d["procesado"] = "no"
                d["observacion"] = ""
                d["inicio"] = ""
                d["fin"] = ""
                d["duracion"] = ""
                print(f"  - {d.get('haber', {}).get('fecha')} | intento: {d['intento']}/{MAX_INTENTOS}")

            with open(path_json, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=4)

            print("JSON actualizado. Se reprocesaran en la siguiente ejecucion.")
            return False

        # Todo procesado sin observaciones -> mover archivos
        fecha = datetime.now().strftime("%d-%m-%Y %H %M %S")

        # Mover Excel
        excels = [f for f in os.listdir(carpeta_proceso) if f.lower().endswith(".xlsx")]
        if excels:
            os.makedirs(carpeta_destino_excel, exist_ok=True)
            nombre_excel = excels[0]
            nuevo_nombre = f"{fecha} _ {nombre_excel}"
            shutil.move(
                os.path.join(carpeta_proceso, nombre_excel),
                os.path.join(carpeta_destino_excel, nuevo_nombre),
            )
            print(f"Excel movido a: {carpeta_destino_excel}/{nuevo_nombre}")

        # Mover JSON (auditoria para devs)
        os.makedirs(carpeta_destino_json, exist_ok=True)
        nuevo_nombre_json = f"{fecha} _ egresos.json"
        shutil.move(
            path_json,
            os.path.join(carpeta_destino_json, nuevo_nombre_json),
        )
        print(f"JSON movido a: {carpeta_destino_json}/{nuevo_nombre_json}")

        return True

    except Exception as e:
        print(f"Error en mover_si_todo_procesado_egresos: {e}")
        return False
