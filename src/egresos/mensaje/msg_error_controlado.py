from typing import List
from src.schema.IEgresos import EgresoItem


def crear_msg_egresos_error_controlado(
    datos: List[EgresoItem] = None,
    titulo="Proceso de egresos detenido de forma controlada"
):
    try:
        if not datos:
            total_registros = 0
            filas = """
            <tr>
                <td colspan="7">No se generaron registros durante esta ejecución.</td>
            </tr>
            """
        else:
            total_registros = len(datos)
            filas = ""

            for d in datos:
                haber = d.get("haber", {})
                filas += f"""
                <tr>
                    <td>{haber.get("fecha", "")}</td>
                    <td>{haber.get("tipo_doc", "")}</td>
                    <td>{haber.get("nro_documento", "")}</td>
                    <td>{haber.get("orden_de", "")}</td>
                    <td>{haber.get("importe", "")}</td>
                    <td>{d.get("observacion", "")}</td>
                    <td>{d.get("procesado", "")}</td>
                </tr>
                """

        html = f"""
        <html>
        <body style="font-family:Arial, Helvetica, sans-serif; font-size:12px;">
            <b>{titulo}</b>
            <br><br>

            Durante la ejecución se presentaron excepciones externas consecutivas que impidieron continuar el flujo con normalidad.
            <br>
            Como medida preventiva, el bot alcanzó el límite máximo de reintentos configurado y detuvo el proceso de forma controlada,
            evitando posibles inconsistencias en los registros.
            <br><br>

            Se procesaron <b>{total_registros}</b> lotes antes de la detención del flujo.
            <br><br>

            <table border="1" cellpadding="4" cellspacing="0" width="100%">
                <tr>
                    <th>Fecha</th>
                    <th>Tip. Doc</th>
                    <th>Nro. Docum.</th>
                    <th>Orden de</th>
                    <th>Importe</th>
                    <th>Observación</th>
                    <th>Procesado</th>
                </tr>
                {filas}
            </table>

            <br>
            El flujo se reanudará automáticamente en la siguiente ejecución, cuando el sistema origen se encuentre más estable.
            <br><br>

            Saludos,<br>
            Sistema
        </body>
        </html>
        """
        return html
    except Exception as e:
        print(f"Error en {__name__}: {e}")
