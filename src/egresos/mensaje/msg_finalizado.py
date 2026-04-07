from typing import List
from src.schema.IEgresos import EgresoItem


def crear_msg_egresos_finalizado(datos: List[EgresoItem] = None, titulo="Proceso de egresos finalizado"):
    try:
        if not datos:
            total_registros = 0
            filas = """
            <tr>
                <td colspan="8">No se encontraron registros.</td>
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
                    <td>{d.get("duracion", "")}</td>
                    <td>{d.get("procesado", "")}</td>
                </tr>
                """

        html = f"""
        <html>
        <body style="font-family:Arial, Helvetica, sans-serif; font-size:12px;">
            <b>{titulo}</b>
            <br><br>
            Se procesaron <b>{total_registros}</b> lotes y terminó la ejecución.
            <br><br>

            <table border="1" cellpadding="4" cellspacing="0" width="100%">
                <tr>
                    <th>Fecha</th>
                    <th>Tip. Doc</th>
                    <th>Nro. Docum.</th>
                    <th>Orden de</th>
                    <th>Importe</th>
                    <th>Observación</th>
                    <th>Duración</th>
                    <th>Procesado</th>
                </tr>
                {filas}
            </table>

            <br>
            Saludos,<br>
            Sistema
        </body>
        </html>
        """
        return html
    except Exception as e:
        print(f"Error en {__name__}: {e}")
