import pyautogui
import time


def detectar_coordenadas_con_rojo():
    print("Mové el mouse. Presiona Ctrl+C para salir")
    
    rojo_min = (150, 0, 0)
    rojo_max = (255, 80, 80)

    try:
        while True:
            x, y = pyautogui.position()
            color = pyautogui.screenshot().getpixel((x, y))  # color actual

            # Comprobar si el color está dentro del rango de rojo
            if rojo_min[0] <= color[0] <= rojo_max[0] and \
               rojo_min[1] <= color[1] <= rojo_max[1] and \
               rojo_min[2] <= color[2] <= rojo_max[2]:
                print(f"¡Fondo rojo detectado en X={x} Y={y}! Color={color}", end="\r")
            else:
                print(f"Coordenadas: X={x} Y={y} Color={color}           ", end="\r")

            time.sleep(0.1)

    except KeyboardInterrupt:
        print("\nDetección finalizada.")


def detectar_coordenadas_con_rojo_cada_segundo():
    print("Mové el mouse. Presiona Ctrl+C para salir")
    
    rojo_min = (150, 0, 0)
    rojo_max = (255, 80, 80)

    try:
        while True:
            x, y = pyautogui.position()
            color = pyautogui.screenshot().getpixel((x, y))

            if rojo_min[0] <= color[0] <= rojo_max[0] and \
               rojo_min[1] <= color[1] <= rojo_max[1] and \
               rojo_min[2] <= color[2] <= rojo_max[2]:
                print(f"[ROJO DETECTADO] Coordenadas X={x} Y={y} Color={color}")
            else:
                print(f"Coordenadas X={x} Y={y} Color={color}")

            time.sleep(1)  # imprime cada 1 segundo

    except KeyboardInterrupt:
        print("\nDetección finalizada.")



if __name__ == "__main__":
    detectar_coordenadas_con_rojo_cada_segundo()
