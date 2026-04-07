def crear_msg_egresos_no_continua(motivo="Motivo no especificado", titulo="Proceso de egresos no continuará"):
    html = f"""
    <html>
    <body style="font-family:Arial, Helvetica, sans-serif; font-size:12px;">
        <b>{titulo}</b>
        <br><br>

        El proceso <b>no continuará</b> debido al siguiente motivo:
        <br><br>

        <b>{motivo}</b>

        <br><br>
        Saludos,<br>
        Sistema
    </body>
    </html>
    """
    return html
