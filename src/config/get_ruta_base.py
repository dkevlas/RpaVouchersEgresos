import sys
import os


def get_ruta_base() -> str:
    if getattr(sys, 'frozen', False):
        path_app = os.path.dirname(sys.executable)
    else:
        path_app = os.getcwd()
    return path_app
