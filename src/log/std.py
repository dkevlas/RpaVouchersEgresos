import sys
from datetime import datetime
from pathlib import Path

class Tee:
    def __init__(self, *files):
        self.files = files

    def write(self, data):
        for f in self.files:
            f.write(data)
            f.flush()

    def flush(self):
        for f in self.files:
            f.flush()


def setup_stdout_logger(base_path: Path):
    """
    Configura el logging para que la salida se muestre en la consola (si existe)
    y siempre se guarde en un archivo de log.
    """
    log_dir = base_path / "log_egreso"
    log_dir.mkdir(exist_ok=True)

    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    log_path = log_dir / f"{timestamp}_bot.log"

    # Abrimos el archivo de log con buffering de línea para escritura inmediata
    log_file = open(log_path, "a", encoding="utf-8", buffering=1)

    if sys.stdout is None:
        # MODO SIN CONSOLA (--noconsole)
        # Redirigimos stdout y stderr directamente al archivo de log.
        sys.stdout = log_file
        sys.stderr = log_file
    else:
        # MODO CON CONSOLA
        # Duplicamos la salida a la consola original y al archivo de log.
        original_stdout = sys.stdout
        original_stderr = sys.stderr
        sys.stdout = Tee(original_stdout, log_file)
        sys.stderr = Tee(original_stderr, log_file)
    
    print(f"--- INICIO DEL LOG ({datetime.now()}) ---")

    return log_file