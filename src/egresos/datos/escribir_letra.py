import time
from pynput.keyboard import Controller
keyboard = Controller()


def escribir_tecla_por_tecla(texto, intervalo=0.3):
    for caracter in texto:
        keyboard.type(caracter) # Esto soporta ñ, tildes y símbolos
        time.sleep(intervalo)