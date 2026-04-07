from datetime import datetime
from typing import List
from src.schema.IEgresos import EgresoItem


def crear_msg_egresos_inicio(
    datos: List[EgresoItem] = None,
    titulo="Inicio de proceso de egresos",
    minutos_por_registro: float = 1.25
):
    try:
        hora_inicio = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

        if not datos:
            total_lotes = 0
            total_registros_global = 0
            tiempo_estimado_min = 0
            filas = """
            <tr>
                <td colspan="5">No hay lotes para procesar.</td>
            </tr>
            """
        else:
            total_lotes = len(datos)
            total_registros_global = 0
            filas = ""

            for d in datos:
                haber = d.get("haber", {})
                suma_deber = round(sum(r.get("importe", 0) for r in d.get("deber", [])), 2)
                cant_haber = 1
                cant_deber = len(d.get("deber", []))
                cant_agregar = len(d.get("agregar", []))
                total_registros = cant_haber + cant_deber + cant_agregar
                total_registros_global += total_registros
                filas += f"""
                <tr>
                    <td>{haber.get("fecha", "")}</td>
                    <td>{haber.get("tipo_doc", "")}</td>
                    <td>{haber.get("nro_documento", "")}</td>
                    <td>{haber.get("orden_de", "")}</td>
                    <td>{suma_deber}</td>
                    <td>{total_registros} (H:{cant_haber} D:{cant_deber} A:{cant_agregar})</td>
                </tr>
                """

        html = f"""
        <html>
        <body style="font-family:Arial, Helvetica, sans-serif; font-size:12px;">
            <b>{titulo}</b>
            <br><br>

            El bot inició la ejecución a las <b>{hora_inicio}</b>.
            <br>
            Se procesarán <b>{total_lotes}</b> lotes con <b>{total_registros_global}</b> registros en total.
            <br>
            Tiempo estimado por registro: <b>{minutos_por_registro} minutos</b>.
            <br>
            Duración total aproximada: <b>{round(total_registros_global * minutos_por_registro, 1)} minutos</b>.
            <br><br>

            <table border="1" cellpadding="4" cellspacing="0" width="100%">
                <tr>
                    <th>Fecha</th>
                    <th>Tip. Doc</th>
                    <th>Nro. Docum.</th>
                    <th>Orden de</th>
                    <th>Importe</th>
                    <th>Registros</th>
                </tr>
                {filas}
            </table>

            <br><br>
            Saludos,<br>
            Sistema
        </body>
        </html>
        """
        return html

    except Exception as e:
        print(f"Error en {__name__}: {e}")
