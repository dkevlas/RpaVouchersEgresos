import subprocess
import re
from datetime import datetime, timedelta
from src.mensaje.enviar_mensaje import enviar_correo_smtp


def get_eventos():
    result = {
        "enviar_mensaje": False,
        "eventos": []  # últimos 3 apagados
    }

    try:
        cmd = [
            "wevtutil", "qe", "System",
            "/q:*[System[(EventID=6006 or EventID=6008 or EventID=1074)]]",
            "/c:3",
            "/rd:true",
            "/f:text"
        ]

        output = subprocess.check_output(cmd, text=True, encoding="utf-8", errors="ignore")

        bloques = output.split("Event[")
        now = datetime.now()
        limite_4h = now - timedelta(hours=4)

        for bloque in bloques:
            if "Event ID:" not in bloque:
                continue

            event_id = int(re.search(r"Event ID:\s*(\d+)", bloque).group(1))
            fecha_evento = re.search(r"Date:\s*(.+)", bloque).group(1)
            fecha_evento = datetime.fromisoformat(fecha_evento.split(".")[0])

            tipo = "NORMAL"
            hora_apagado_real = fecha_evento

            if event_id == 6008:
                tipo = "INESPERADO"
                m = re.search(r"a las (\d+:\d+:\d+)", bloque)
                if m:
                    hora_apagado_real = datetime.combine(
                        fecha_evento.date(),
                        datetime.strptime(m.group(1), "%H:%M:%S").time()
                    )

            evento = {
                "event_id": event_id,
                "tipo": tipo,
                "fecha_registro": fecha_evento,
                "hora_apagado": hora_apagado_real
            }

            result["eventos"].append(evento)

            # ¿Apagado inesperado en últimas 4 horas?
            if tipo == "INESPERADO" and hora_apagado_real >= limite_4h:
                result["enviar_mensaje"] = True

        return result

    except Exception as e:
        return {
            "enviar_mensaje": False,
            "error": str(e),
            "eventos": []
        }


def build_html_mensaje(data):
    apagados = [e for e in data["eventos"] if e["tipo"] == "INESPERADO"]

    if not apagados:
        return ""

    rows = ""
    for e in apagados:
        rows += f"""
        <tr>
            <td>{e['hora_apagado'].strftime('%Y-%m-%d %H:%M:%S')}</td>
            <td>{e['event_id']}</td>
            <td>{e['tipo']}</td>
        </tr>
        """

    html = f"""
    <html>
        <body>
            <h3>⚠ Apagados inesperados detectados</h3>
            <p>Se detectaron apagados inesperados en las últimas horas.</p>
            <table border="1" cellpadding="6" cellspacing="0">
                <tr>
                    <th>Hora de apagado</th>
                    <th>Event ID</th>
                    <th>Tipo</th>
                </tr>
                {rows}
            </table>
        </body>
    </html>
    """

    return html


def mensaje_eventos():
    try:
        result = get_eventos()
        
        body = build_html_mensaje(data=result)
        
        if result.get('enviar_mensaje'):
            print("SE ENVIARÁ MENSAJE")
            enviar_correo_smtp(
                asunto="Apagado inesperado del servidor (últimas 4 horas)",
                body=body
            )

    except Exception as e:
        print(f"Error insperado con la información de los eventos {e}")
    
    
if __name__ == '__main__':
    mensaje_eventos()
