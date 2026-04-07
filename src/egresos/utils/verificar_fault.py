import time
import pyautogui
from src.egresos.excel.escritor_egresos import EscritorEgresos
from src.sistema.fault import buscar_error_siep, cerrar_ventana_fault
from src.egresos.debug.screenshow import ScreenImgEgresos


def verificar_fault(lote, write_log: EscritorEgresos):
    try:
        print("[Aviso] Verificando errores...")
        pyautogui.click(425, 8)
        time.sleep(2)   
        
        error_id = buscar_error_siep()

        if error_id:
            print(f"¡Ventana detectada! ID: {error_id}. Procediendo a cerrar...")
            
            ScreenImgEgresos.here(lote=lote, nombre=f"siep_no_responde")
            
            # 2. Cerrar si fue detectado
            if cerrar_ventana_fault(error_id):
                print("Orden de cierre enviada con éxito.")
            else:
                print("Error al intentar enviar la orden de cierre.")
                    
            write_log.escribir_observacion(lote, "Error por inestabilidad del sistema (No responde el app)")
            time.sleep(5)
            return "ERROR"
        elif error_id is False:
            print("Estado: Todo normal, no hay errores visibles.")
            return True
        else:
            print("Estado: Fallo preventivo en el escaneo.")
            return None
    except Exception as e:
        print(f"Error en verificar_fault: {e}")
        return None
