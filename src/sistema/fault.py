import win32gui
import win32con
import pyautogui
import time


def buscar_error_siep():
    """Busca la ventana y retorna su HWND (ID) si existe, de lo contrario False o None."""
    resultado = {"hwnd": None}

    def callback(hwnd, extra):
        if win32gui.IsWindowVisible(hwnd):
            titulo = win32gui.GetWindowText(hwnd).lower()
            # Criterios de búsqueda
            if "dejó de funcionar" in titulo or "siep20.exe" in titulo:
                resultado["hwnd"] = hwnd
                return False  # Detener búsqueda al encontrarla
        return True

    try:
        win32gui.EnumWindows(callback, None)
        if resultado["hwnd"]:
            return resultado["hwnd"] # Retorna el ID de la ventana
        return False # No se encontró
    except Exception:
        return None # Error preventivo

def cerrar_ventana_fault(hwnd):
    """Recibe un HWND y le ordena a Windows cerrar esa ventana específica."""
    try:
        # Enviamos el mensaje de cierre (Equivale a presionar la X)
        win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
        return True
        print("CIERRE")
    except Exception as e:
        print(f"No se pudo cerrar la ventana: {e}")
        return False


def cierre_fast():
    try:
        error_id = buscar_error_siep()

        if error_id:
            print(f"¡Ventana detectada! ID: {error_id}. Procediendo a cerrar...")
            
            # 2. Cerrar si fue detectado
            if cerrar_ventana_fault(error_id):
                print("Orden de cierre enviada con éxito.")
                for _ in range(2):
                    pyautogui.click(445, 344)
                    time.sleep(0.1)

                time.sleep(5)
            else:
                print("Error al intentar enviar la orden de cierre.")
                
        elif error_id is False:
            print("Estado: Todo normal, no hay errores visibles.")
        else:
            print("Estado: Fallo preventivo en el escaneo.")
    except Exception as e:
        print(f"[ERROR_CIERRE_FAST] {e}")
