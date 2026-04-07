import subprocess
import time


def cerrar_siep_seguro():
    nombre_exe = "siep20.exe"
    try:
        # Verificar si el proceso existe antes de matarlo
        check = subprocess.run(
            f'tasklist /FI "IMAGENAME eq {nombre_exe}" /NH',
            shell=True,
            capture_output=True,
            text=True
        )

        if nombre_exe.lower() in check.stdout.lower():
            # El proceso existe, cerrarlo
            resultado = subprocess.call(
                f'taskkill /f /im {nombre_exe}',
                shell=True,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL
            )
            if resultado == 0:
                print(f"SIEP cerrado correctamente al finalizar el bot.")
                time.sleep(2)
                return True
            else:
                print(f"No se pudo cerrar {nombre_exe}.")
                return False
        else:
            print(f"SIEP ya no estaba abierto (posible cierre inesperado durante la ejecucion).")
            return True

    except Exception as e:
        print(f"Ocurrio un error al intentar cerrar: {e}")
        return False


if __name__ == "__main__":
    cerrar_siep_seguro()
