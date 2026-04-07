"""
Ejecutar: python leer_celda.py
Tienes 5 segundos para posicionar el cursor en la celda.
"""
import time
import pyautogui
import pyperclip


class UtilsModificar:
    @staticmethod
    def leer_celda():
        print("Posiciona el cursor en la celda...")
        print("Leyendo en 5 segundos...")
        time.sleep(5)
        pyperclip.copy("")
        pyautogui.hotkey('ctrl', 'c')
        time.sleep(0.5)
        valor = pyperclip.paste().strip()
        if valor:
            print(f"HAY VALOR: '{valor}'")
            return True
        else:
            print("NO HAY VALOR (vacio)")
            time.sleep(3)
            return False
    
    @staticmethod
    def quitar_fila():
        print("Quitar fila...")
        time.sleep(1)
        pyautogui.click(731, 145)
        print("Fila quitada.")
        time.sleep(5)
        return True
    
    @staticmethod
    def nuevamente_click_en_cuenta():
        print("Nuevamente click en cuenta de primera fila...")
        time.sleep(1)
        pyautogui.click(255, 195)
        print("Click en cuenta.")
        time.sleep(5)
        return True

    @staticmethod
    def click_agregar_nuevamente():
        print("Nuevamente click en agregar...")
        time.sleep(1)
        pyautogui.click(671, 145)
        print("Click en agregar.")
        time.sleep(5)
        return True

    @staticmethod
    def reposionarse_a_ultima_fila(cant: int = 20):
        print("Reposionarse a ultima fila...")
        time.sleep(1)
        pyautogui.click(63, 195)
        print("Click en ultima fila.")
        time.sleep(5)
        for i in range(cant):
            pyautogui.press('down')
            time.sleep(0.5)
        return True
    
    @staticmethod
    def click_guardar():
        print("Guardar para reintentar")
        for _ in range(2):
            pyautogui.click(196, 145)
            time.sleep(1)
        time.sleep(2)

    @staticmethod
    def click_modificar(intento=2):
        print("MOdificar una vez mas")
        for _ in range(intento):
            pyautogui.click(145, 145)
            time.sleep(1)
        if intento == 1:
            time.sleep(5)
            return
        time.sleep(2)


    @staticmethod
    def click_cerrar_nieto():
        print("Apunto de cerrar la ventana nieta")
        time.sleep(2)
        for _ in range(2):
            pyautogui.click(1011, 24)
            time.sleep(1)
        time.sleep(5)


    def buscar_voucher():
        print("Buscanco el último voucher")
        pyautogui.click(498, 244)
        time.sleep(4)
        pyautogui.click(141, 274)
        time.sleep(20)
        pyautogui.click(882, 214)
        time.sleep(3)


if __name__ == "__main__":
    UtilsModificar.leer_celda()
