import os
import sys
import time
import traceback
from pathlib import Path
from src.config.path_folders import PathFolder
from src.log.std import setup_stdout_logger
from bot_egresos import botVouchersEgresos


base_path = Path(sys.executable).parent

setup_stdout_logger(base_path)

LOCK_FILE = os.path.join(PathFolder.FOLDER_BASE, 'app.lock')

ELIMINAR = True

try:
    botVouchersEgresos(cant=3)
except KeyboardInterrupt as k:
    print(f'\n\n[TECLADO] interrupcion {k}\n')
    time.sleep(10)
except Exception as e:
    print(f'\n\n[ERROR GLOBAL]\n\n{e}\n\n')
    traceback.print_exc()
    time.sleep(10)
finally:
    try:
        print(f"Intentando eliminar el archivo lock: {LOCK_FILE} - {ELIMINAR}")

    except Exception as e:
        print(f"Error al cerrar los logs o el app.lock???: {e}")
