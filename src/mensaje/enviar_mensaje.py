import os
import win32com.client
from pathlib import Path
from src.config.path_folders import PathFolder
import smtplib
from email.message import EmailMessage


def body():
    return """
<html>
<body style="font-family: Arial, Helvetica, sans-serif; font-size: 13px; color: #000;">
    <p>Estimados,</p>

    <p>
        El proceso de <b>registro de información</b> en el ERP ha finalizado correctamente.
    </p>

    <p>
        Los datos fueron leídos desde el archivo Excel y registrados de forma secuencial
        según la cantidad de filas procesadas.
    </p>

    <p>
        De no presentarse observaciones, no se requiere ninguna acción adicional.
    </p>

    <br>

    <p style="font-size: 12px; color: #555;">
        Este correo fue generado automáticamente por el sistema.<br>
        Por favor, no responder.
    </p>

    <br>

    <p>
        Atentamente,<br>
        <b>Sistema de Registro ERP</b>
    </p>
</body>
</html>
"""


def enviar_correo_outlook() -> None:
    try:
        correo_destino = "testing@sertech.pe"
        
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)

        mail.To = correo_destino
        mail.Subject = "Proceso de registro ERP finalizado"
        mail.HTMLBody = body()

        adjuntos = [f for f in os.listdir(PathFolder.FOLDER_BASE) if f.lower().endswith((".xls", ".xlsx"))]
        
        if adjuntos:
            for archivo in adjuntos:
                full_archivo = os.path.join(PathFolder.FOLDER_BASE, archivo)
                mail.Attachments.Add(str(Path(full_archivo)))

        mail.Send()
    except Exception as e:
        print(f"[ERROR OUTLOOK] - {e}")


PASSWORD = '123456'

def enviar_correo_smtp(body: str = None, asunto: str = None):
    try:
        #return      
        msg = EmailMessage()
        msg["From"] = "robot-procasa@fastcloud.pe"
        msg["To"] = "procasa.eliasaguirre@gmail.com"
        #msg["To"] = "dennis.blas@sertech.pe"
        msg["Subject"] = "Proceso de registro ERP" if not asunto else asunto
        msg["Bcc"] = "dennis.blas@sertech.pe" # con coma se separa
        msg.set_content("El proceso de registro ERP ha finalizado")
        msg.add_alternative(body, subtype="html")

        for f in os.listdir(PathFolder.FOLDER_PROCESO):
            if f.lower().endswith(".xlsx"):
                ruta = os.path.join(PathFolder.FOLDER_PROCESO, f)
                with open(ruta, "rb") as file:
                    msg.add_attachment(
                        file.read(),
                        maintype="application",
                        subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        filename=f
                    )

        with smtplib.SMTP("mail.fastcloud.pe", 2525, timeout=20) as server:
            server.set_debuglevel(0)
            server.login("robot-procasa@fastcloud.pe", PASSWORD)
            server.send_message(msg)
        
        # with smtplib.SMTP_SSL("mi3-sr23.supercp.com", 465, timeout=20) as server:
        #     server.set_debuglevel(0)   #
        #     server.login("robot-procasa@fastcloud.pe", PASSWORD)
        #     server.send_message(msg)

    except Exception as e:
        print(f"[ERROR SMTP] - {e}")


if __name__ == "__main__":
    enviar_correo_smtp()
