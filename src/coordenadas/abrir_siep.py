import subprocess
import pyautogui
import time


NOMBRE_EXE = "siep20.exe"
RUTA_SIEP = r"D:\SI_PROCASA\siep.exe"


def siep_esta_abierto():
    try:
        check = subprocess.run(
            f'tasklist /FI "IMAGENAME eq {NOMBRE_EXE}" /NH',
            shell=True,
            capture_output=True,
            text=True
        )
        return NOMBRE_EXE.lower() in check.stdout.lower()
    except Exception:
        return False


def abrir_siep_por_coordenadas(max_intentos=3):
    try:
        for intento in range(1, max_intentos + 1):
            print(f"Intento {intento}/{max_intentos} de abrir SIEP...")

            if siep_esta_abierto():
                print("SIEP ya esta abierto.")
                return True
            
            # Cerrar cualquier capa/overlay invisible que bloquee el escritorio
            pyautogui.press('escape')
            time.sleep(0.5)
            pyautogui.hotkey('alt', 'f4')
            time.sleep(1)
            pyautogui.press('escape')
            time.sleep(0.5)

            print("Minimizando ventanas (Win + D)...")
            pyautogui.hotkey('win', 'd')
            time.sleep(2)

            # --- MÉTODO ROBUSTO USANDO EL MENÚ DE INICIO ---
            print("Abriendo SIEP a través del menú de inicio...")
            pyautogui.press('win')
            time.sleep(1)  # Esperar a que el menú de inicio aparezca

            pyautogui.write('SIEP', interval=0.1) # Escribir el nombre del programa
            time.sleep(1.5) # Esperar a que la búsqueda finalice

            pyautogui.press('enter')
            time.sleep(12) # Dar tiempo de sobra para que la aplicación cargue
            # --- FIN DEL MÉTODO ROBUSTO ---

            if siep_esta_abierto():
                print("SIEP abierto correctamente.")
                return True

            print(f"SIEP no se detecto como abierto en intento {intento}.")
            # Si falla, cerramos cualquier ventana extra que se haya abierto
            pyautogui.hotkey('alt', 'f4')
            time.sleep(2)

        # Último recurso: abrir SIEP desde ruta para romper la capa invisible, cerrarlo, y reintentar con pyautogui
        print("Intentando romper capa invisible abriendo SIEP desde subprocess...")
        try:
            subprocess.Popen(RUTA_SIEP, shell=True)
            time.sleep(10)
            # Cerrar el SIEP que abrimos (no sirve para el flujo)
            subprocess.call(f'taskkill /f /im {NOMBRE_EXE}', shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            time.sleep(5)
            print("SIEP cerrado. Reintentando apertura normal...")
        except Exception as e:
            print(f"Error en subprocess fallback: {e}")

        # Reintentar una vez más con el método normal
        pyautogui.hotkey('win', 'd')
        time.sleep(2)
        pyautogui.press('win')
        time.sleep(1)
        pyautogui.write('SIEP', interval=0.1)
        time.sleep(1.5)
        pyautogui.press('enter')
        time.sleep(12)

        if siep_esta_abierto():
            print("SIEP abierto correctamente tras romper capa invisible.")
            return True

        print("ERROR: No se pudo abrir SIEP despues de todos los intentos.")
        return "NO_OPEN"
    except Exception as e:
        print(f"Error al abrir el programa: {e}")
        return False


def _abrir_siep_por_coordenadas_(max_intentos=3):
    icono_x, icono_y = 139, 388
    try:
        for intento in range(1, max_intentos + 1):
            print(f"Intento {intento}/{max_intentos} de abrir SIEP...")

            if siep_esta_abierto():
                print("SIEP ya esta abierto.")
                return True

            print("Minimizando ventanas (Win + D)...")
            pyautogui.hotkey('win', 'd')

            pyautogui.click(559, 348)
            
            time.sleep(2.5)

            pyautogui.doubleClick(icono_x, icono_y)
            time.sleep(12)

            if siep_esta_abierto():
                print("SIEP abierto correctamente.")
                return True

            print(f"SIEP no se detecto como abierto en intento {intento}.")
            time.sleep(5)

        print("ERROR: No se pudo abrir SIEP despues de todos los intentos.")
        return "NO_OPEN"
    except Exception as e:
        print(f"Error al abrir el programa: {e}")
        return False


if __name__ == "__main__":
    abrir_siep_por_coordenadas()
