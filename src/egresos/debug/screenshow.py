import os
import shutil
import pyautogui
from pathlib import Path
from src.egresos.rutas.config import ConfigRutas
from src.schema.IEgresos import EgresoItem
from datetime import datetime


class ScreenImgEgresos:
    @staticmethod
    def _obtener_ini():
        archivos = [f for f in os.listdir(ConfigRutas.FOLDER_PROCESO) if f.lower().endswith('.xlsx')]
        if not archivos:
            return "GEN"
        solo_file = archivos[0]
        if solo_file.upper().startswith("M"):
            return "MEE"
        elif solo_file.upper().startswith("H"):
            return "HDE"
        elif solo_file.upper().startswith("E"):
            return "ELE"
        return "GEN"

    @staticmethod
    def mover_intentos_anteriores(lote: EgresoItem):
        try:
            fecha_lote = lote['haber']['fecha']
            partes = fecha_lote.split("/")
            dir_mes = f"{partes[1]}-{partes[2]}"
            dir_fecha = fecha_lote.replace("/", "-")
            ini = ScreenImgEgresos._obtener_ini()

            dir_final = Path(ConfigRutas.FOLDER_IMG) / ini / dir_mes / dir_fecha

            if not dir_final.exists():
                return

            imagenes = [f for f in os.listdir(dir_final) if f.lower().endswith(('.jpg', '.png'))]
            if not imagenes:
                return

            # Buscar siguiente número de intento
            n = 1
            while (dir_final / str(n)).exists():
                n += 1

            carpeta_intento = dir_final / str(n)
            os.makedirs(carpeta_intento)

            for img in imagenes:
                shutil.move(str(dir_final / img), str(carpeta_intento / img))

            print(f"Imagenes anteriores movidas a intento {n}: {carpeta_intento}")
        except Exception as e:
            print(f"Error al mover imagenes de intentos anteriores: {e}")

    @staticmethod
    def here(lote: EgresoItem, nombre: str = None, haber: bool = False, deber: bool = False, add: bool = False, num: int = 0):
        try:

            fecha_lote = lote['haber']['fecha']  # "dd/MM/yyyy"
            partes = fecha_lote.split("/")       # ["dd", "MM", "yyyy"]
            dir_mes = f"{partes[1]}-{partes[2]}" # "MM-yyyy"
            dir_fecha = fecha_lote.replace("/", "-")  # "dd-MM-yyyy"

            ini = ScreenImgEgresos._obtener_ini()

            fecha = datetime.now().strftime("%d-%m-%Y %H %M %S")

            if num + 1 > 1:
                nombre = f"{num + 1}_{nombre}"

            dir_final = Path(ConfigRutas.FOLDER_IMG) / ini / dir_mes / dir_fecha
            
            if haber:
                filename = Path(dir_final) / f'{nombre}_{lote["haber"]["nro_documento"]}_haber.jpg'
            elif deber:
                filename = Path(dir_final) / f'{nombre}_deber.jpg'
            elif add:
                filename = Path(dir_final) / f'{nombre}_add.jpg'
            else:
                filename = Path(dir_final) / f'{nombre}_{fecha}.jpg'

            os.makedirs(dir_final, exist_ok=True)
            
            filename = Path(dir_final) / f'{nombre}_{fecha}.jpg'
            
            if nombre:
                filename = os.path.join(dir_final, f'{nombre}_{fecha}.jpg')

            screenshot = pyautogui.screenshot()
            screenshot.save(filename)
        except Exception as e:
            print(f'Error al tomar captura de pantalla {e}')
