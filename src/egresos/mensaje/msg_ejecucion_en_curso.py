from datetime import datetime


def crear_msg_egresos_ejecucion_en_curso(
    titulo="Proceso de egresos detenido",
    motivo="Se detectó una ejecución en curso para evitar conflictos."
):
    try:
        hora_evento = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

        html = f"""
        <html>
        <body style="font-family:Arial, Helvetica, sans-serif; font-size:12px;">
            <b>{titulo}</b>
            <br><br>

            El bot inició la ejecución a las <b>{hora_evento}</b>, sin embargo,
            el proceso será <b>detenido</b> debido a que ya existe una ejecución
            en curso.
            <br><br>

            Motivo:
            <br>
            <b>{motivo}</b>
            <br><br>

            Esta validación se realiza para evitar conflictos de información
            y ejecuciones simultáneas.
            <br><br>

            Saludos,<br>
            Sistema
        </body>
        </html>
        """
        return html

    except Exception as e:
        print(f"Error en {__name__}: {e}")
