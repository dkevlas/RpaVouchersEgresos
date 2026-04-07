import time
import pyautogui
import ctypes

class LoginCoordenadas():
    
    @staticmethod
    def asegurar_capslock_apagado():
        VK_CAPITAL = 0x14
        # Si Caps Lock está encendido, lo apaga
        if ctypes.windll.user32.GetKeyState(VK_CAPITAL) & 1:
            pyautogui.press('capslock')
            time.sleep(0.1)

    @staticmethod
    def escribir_usuario_contraseña_click():
        usuario_coord = (628, 389)
        contraseña_coord = (633, 445)
        click_final_coord = (200, 503)

        usuario_texto = "sertech"
        contraseña_texto = "12345678"

        print("Coloca la ventana correcta. Iniciando en 2 segundos...")
        time.sleep(2)

        # --- Asegurar Caps Lock apagado ---
        LoginCoordenadas.asegurar_capslock_apagado()

        # --- Usuario ---
        pyautogui.click(usuario_coord)
        time.sleep(0.1)
        for _ in range(10):
            pyautogui.press('backspace')
        for _ in range(10):
            pyautogui.press('delete')

        LoginCoordenadas.asegurar_capslock_apagado()
        pyautogui.typewrite(usuario_texto, interval=0.05)
        time.sleep(0.2)

        # --- Contraseña ---
        pyautogui.click(contraseña_coord)
        time.sleep(0.1)
        for _ in range(10):
            pyautogui.press('backspace')
        for _ in range(10):
            pyautogui.press('delete')

        LoginCoordenadas.asegurar_capslock_apagado()
        pyautogui.typewrite(contraseña_texto, interval=0.05)
        time.sleep(0.5)

        # --- Enter ---
        pyautogui.press('enter')
        print("ENTER enviado...")
        time.sleep(1.5)

        # --- Click final ---
        pyautogui.click(click_final_coord)
        print("Usuario y contraseña escritos, click final realizado.")
        time.sleep(5)


    @staticmethod
    def seleccionar_sucursal():
        coord = (440, 387)
        print(f"Haciendo doble clic en X={coord[0]} Y={coord[1]} en 2 segundos...")
        time.sleep(1)
        pyautogui.doubleClick(coord)
        print("Doble clic realizado.")
        time.sleep(2)
        
        # --- Cambiar contraseña ---
        coord_temp = (694, 503)
        pyautogui.click(coord_temp)
