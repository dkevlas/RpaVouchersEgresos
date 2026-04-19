import pyautogui
import pyperclip
import time
import threading
from pywinauto import Application, mouse
from src.utils.safe_type import safe_type
from src.egresos.ot.ot import obtener_dni_cta_ot
from src.egresos.datos.escribir_letra import escribir_tecla_por_tecla
from src.egresos.debug.screenshow import ScreenImgEgresos
from src.egresos.debug.leer_celda import UtilsModificar
from src.egresos.ot.mapeoRuc import obtener_ruc_si_especial


class BotEgresos():
    def __init__(self):
        self.app = None
        self.voucher = None
    
    
    def connect(self):
        self.app = Application(backend="win32").connect(
            title_re=".*PRODUCTORES Y COMERCIANTES.*",
            timeout=10
        )
        
    def openVoucherEgresos(self):
        try:
            if self.app is None:
                self.connect()

            # ⚠️ usa la ventana principal por título, NO top_window
            dlg = self.app.window(title_re=".*PRODUCTORES Y COMERCIANTES.*")
            dlg.set_focus()
            time.sleep(0.5)

            dlg.type_keys("%f")
            time.sleep(0.5)

            dlg.type_keys("c")
            time.sleep(0.3)

            dlg.type_keys("{DOWN}")
            time.sleep(0.5)
            dlg.type_keys("{DOWN}")
            time.sleep(0.5)
            
            dlg.type_keys("{ENTER}")

        except Exception as e:
            print(f"[openVoucherEgresos] {repr(e)}")
            raise

    def focus_ventana_voucher(self):
        try:
            main_dlg = self.app.top_window()
            v_voucher = main_dlg.child_window(
                title="PRODUCTORES Y COMERCIANTES ASOCIADOS S .R.L.-Voucher de Egresos",
                top_level_only=False
            )
            self.voucher = v_voucher
            v_voucher.set_focus()
            print("¡Ventana de Voucher enfocada!")

        except Exception as e:
            print(f"[focus_ventana_voucher]: {e}")

    def llenar_cuenta(self, ini: str):
        try:
            
            if ini.startswith("M"):
                cuenta = "CAJA TDA. MEIGGS - CHIMBOTE"
            elif ini.startswith("E"):
                cuenta = "CAJA TDA. ELIAS AGUIRRE -CHIMBOTE"
            elif ini.startswith("H"):
                cuenta = "CAJA TDA. HIDRAULICA - CHIMBOTE"

            
            v = self.voucher
            edits = v.descendants(class_name="Edit")
            
            input_cuenta = edits[2]
            input_cuenta.click_input()
            input_cuenta.set_focus()
            input_cuenta.type_keys(cuenta, with_spaces=True)
            input_cuenta.type_keys("{ENTER}")

            print("Campo Cuenta llenado")

        except Exception as e:
            print(f"[llenar_cuenta] {e}")

    def llenar_fecha(self, fecha: str):
        v = self.voucher
        
        v.type_keys("{TAB}")
        time.sleep(0.2)

        v.type_keys("{DELETE}" * 8)
        time.sleep(0.1)
        v.type_keys("{HOME}")

        time.sleep(0.2)
        
        digitos = fecha.replace("/", "").replace("-", "")
        
        for d in digitos:
            v.type_keys(d)
            time.sleep(0.05)

        v.type_keys("{TAB}")
        time.sleep(0.2)

        print(f"Fecha enviada correctamente: {fecha}")

        try:
            dlg = self.app.window(title_re=".*Atención.*")
            dlg.set_focus()

            dlg.type_keys("{RIGHT}")
            time.sleep(0.1)
            dlg.type_keys("{ENTER}")
            time.sleep(0.1)
        except:
            pass

        v.type_keys("{TAB}")
        time.sleep(0.2)
        
        print("Fecha y confirmación OK")

    def actividad(self):
        try:
            v = self.voucher
            v.type_keys("{F1}")
            time.sleep(0.5)
            
            dlg = self.app.window(title_re=".*Lista de Actividades.*")
            dlg.set_focus()
            time.sleep(0.3)
            
            dlg.type_keys("1")
            time.sleep(1)
            dlg.type_keys("{ENTER}{ENTER}")
            time.sleep(1)
            v.type_keys("{TAB}")
            print("Actividad seleccionado")
        except Exception as e:
            print(f"[actividad] {e}")
    
    def flujo(self):
        try:
            v = self.voucher
            v.type_keys("{F1}")
            time.sleep(0.5)
            
            dlg = self.app.window(title_re=".*Lista General Flujos*")
            dlg.set_focus()
            time.sleep(0.3)
            
            dlg.type_keys("5")
            time.sleep(1)
            dlg.type_keys("{ENTER}{ENTER}") 
            time.sleep(1)
            v.type_keys("{TAB}")
            print("Flujo Pago a proveedores de bienes y servicios seleccionado seleccionado")
            time.sleep(1)
        except Exception as e:
            print(f"[flujo] {e}")
    
    def seleccionar_documento(self):
        try:
            v = self.voucher
            v.type_keys("{DOWN}")
            time.sleep(0.3)
            v.type_keys("{TAB}")
            print("Efectivo seleccionado")
            time.sleep(1)
        except Exception as e:
            print(f"[documento] {e}")
    
    def numero_documento(self, numero_documento: str):
        try:
            v = self.voucher
            v.type_keys(numero_documento, with_spaces=True)
            time.sleep(1)
            v.type_keys("{TAB}")
            time.sleep(1)
            #v.type_keys("{TAB}")
            #time.sleep(1)
            print("Numero de documento escrito")
        except Exception as e:
            print(f"[numero_documento] {e}")

    def ir_a_favor_de_control(self, cliente: str, ot: bool):
        try:
            v = self.voucher
            ventana_rect = self.voucher.rectangle()
            
            offset_x = 351 - ventana_rect.left
            offset_y = 416 - ventana_rect.top
            
            click_x = ventana_rect.left + offset_x
            click_y = ventana_rect.top + offset_y
            
            self.voucher.set_focus()
            time.sleep(0.2)
            pyautogui.click(click_x, click_y)
            time.sleep(0.3)
            pyautogui.press('f1')
            print(f"F1 enviado en coordenadas dinámicas ({click_x}, {click_y})")

            time.sleep(3)

            dlg = self.app.window(title_re=".*Lista de Cliente Proveedores.*")
            dlg.set_focus()

            if ot:
                print(f"[debug] Buscando cliente: {cliente}")
                
                dni, cta = obtener_dni_cta_ot(nombre=cliente)
                
                print(f"[debug] DNI: {dni}, CTA: {cta}")
                
                dni = dni.upper()
                time.sleep(0.5)

                print("Limpiando el campo")
                safe_type(dlg, "^a{DEL}")

                if not safe_type(dlg, dni, timeout=12, with_spaces=True):
                    return False
            else:
                ruc = obtener_ruc_si_especial(cliente)
                if ruc:
                    # Buscar por RUC en vez del nombre (tiene caracteres especiales)
                    print(f"[depositario] Usando RUC '{ruc}' en lugar de '{cliente}'")
                    if not safe_type(dlg, ruc, timeout=12, with_spaces=True):
                        return False
                else:
                    for key in ["{TAB}", "{TAB}", "{UP}", "{TAB}", "{TAB}", "{TAB}", "{TAB}"]:
                        if not safe_type(dlg, key, timeout=300):
                            print("[depositario] La ventana no respondió al enviar teclas")
                            return False

                    if not safe_type(dlg, cliente, timeout=12, with_spaces=True):
                        return False

            time.sleep(1)
            if cliente.strip() == 'VARIOS':
                print('VARIOS??')
                pyautogui.click(327, 321)
                time.sleep(2)
            
            safe_type(dlg, "{ENTER}{ENTER}")
            v.type_keys("{TAB}")

            print("Depositario seleccionado")

            try:
                dlg.close()
                print("[debug] Ventana cerrada: cliente NO existe")
                msg = f"Cliente {cliente} NO existe"
                self.cerrar_voucher(p=0)
                return f"CONTROL-{msg}"
            except Exception as e:
                print("[debug] Ventana no se pudo cerrar: cliente EXISTE")
                return True

        except Exception as e:
            print(f"Fallo: {e}")
            self.voucher.set_focus()
            pyautogui.click(351, 416)
            time.sleep(0.3)
            pyautogui.press('f1')
            return "MODO_COORDENADAS"
    
    def obtener_valor_amarillo_haber(self):
        try:
            # Forzamos la búsqueda asegurando que el elemento sea visible/exista
            campo = self.voucher.child_window(
                auto_id="20", 
                control_type="Edit"
            ).wrapper_object() # wrapper_object ayuda a acceder a métodos específicos

            # Intentamos obtener el valor de tres formas por orden de fiabilidad:
            texto = campo.get_value() or campo.window_text() or campo.legacy_properties().get('Value', '')

            print(f"Valor obtenido del Edit (auto_id=20): {texto}")
            return texto.strip() if texto else None

        except Exception as e:
            # El error en tu imagen muestra que 'e' es un objeto con mucha info, 
            # mejor imprimir solo el mensaje si es posible o el tipo.
            print(f"Error al leer Edit 20: {type(e).__name__} - {e}")
            return None
    
    def ingresar_monto(self, monto: float):
        try:
            import pyautogui
            monto_str = str(monto)
            
            self.voucher.set_focus()
            time.sleep(0.5)
            
            rect_ventana = self.voucher.rectangle()

            click_x = 574
            click_y = 438
            
            print(f"Haciendo clic en coordenadas de control: ({click_x}, {click_y})")
            
            # 4. Ejecutar acciones físicas
            pyautogui.click(click_x, click_y)
            time.sleep(0.3)
            
            # Doble clic opcional para asegurar que el cursor esté dentro
            pyautogui.click(click_x, click_y) 
            
            # Limpiar campo (Ctrl+A -> Backspace)
            pyautogui.hotkey('ctrl', 'a')
            pyautogui.press('backspace')
            time.sleep(0.2)
            
            # Escribir el monto
            pyautogui.write(monto_str)
            time.sleep(0.2)
            
            # Confirmar con ENTER o TAB
            pyautogui.press('enter')
            
            print(f"✅ Proceso de ingreso completado para monto: {monto_str}")
            
        except Exception as e:
            print(f"❌ [ingresar_monto] Error por coordenadas: {e}")

    def llenar_glosa(self, glosa: str):
        try:
            v = self.voucher

            for _ in range(3):
                v.type_keys("{TAB}")
                time.sleep(0.5)

            v.type_keys(glosa, with_spaces=True)
            time.sleep(0.2)
            v.type_keys("{ENTER}")
            for i in range(4):
                time.sleep(0.2)
                v.type_keys("{TAB}")
                
            time.sleep(0.2)
            print("Campo Glosa llenado")
            return True
        except Exception as e:
            print(f"[llenar_glosa]: {e}")
    
    def obtener_valor_campo_cargos(self):
        try:
            campo = self.voucher.child_window(control_id=1012, class_name="PBEDIT105")
            texto = campo.window_text().strip()

            print(f"Valor obtenido de cargo: {texto}")
            return float(texto.replace(",", ""))
        except (ValueError, TypeError, Exception):
            return None

    def obtener_valor_campo_abonos(self):
        try:
            campo = self.voucher.child_window(control_id=1009, class_name="PBEDIT105")
            texto = campo.window_text().strip()

            print(f"Valor obtenido de abono: {texto}")
            return float(texto.replace(",", ""))
        except (ValueError, TypeError, Exception) as e:
            print(f"Error Abonos: {e}")
            return None
    
    def _ventana_sigue_abierta(self):
        try:
            return self.voucher.exists() and self.voucher.is_visible()
        except Exception:
            return False

    def _esperar_que_responda(self, titulo=".*PRODUCTORES.*", timeout=120, intervalo=10, lote=None):
        """Espera hasta que la ventana responda (no esté congelada)."""
        import subprocess
        inicio = time.time()
        intento = 0
        screenshot_tomado = False

        while (time.time() - inicio) < timeout:
            intento += 1

            # En el primer intento, mover la ventana 0px para forzar "(No responde)" y tomar pantallazo
            if intento == 1 and not screenshot_tomado:
                try:
                    import ctypes
                    import win32gui
                    hwnd = win32gui.FindWindow(None, None)
                    # Buscar ventana de SIEP por titulo parcial
                    def find_siep(hwnd, _):
                        if win32gui.IsWindowVisible(hwnd):
                            t = win32gui.GetWindowText(hwnd)
                            if "PRODUCTORES" in t.upper():
                                find_siep.handle = hwnd
                        return True
                    find_siep.handle = None
                    win32gui.EnumWindows(find_siep, None)

                    if find_siep.handle:
                        # SendMessage WM_NULL fuerza deteccion de "No responde" sin hacer nada
                        ctypes.windll.user32.SendMessageTimeoutW(
                            find_siep.handle, 0x0000, 0, 0, 0x0002, 3000, ctypes.byref(ctypes.c_ulong())
                        )
                    time.sleep(2)
                    if lote:
                        ScreenImgEgresos.here(lote=lote, nombre="programa_no_responde")
                    screenshot_tomado = True
                    print("Screenshot tomado: programa no responde.")
                except Exception:
                    pass

            try:
                app = Application(backend="uia").connect(title_re=titulo, timeout=5)
                win = app.window(title_re=titulo)
                resultado = [False]

                def _check():
                    try:
                        win.window_text()
                        resultado[0] = True
                    except Exception:
                        pass

                hilo = threading.Thread(target=_check, daemon=True)
                hilo.start()
                hilo.join(timeout=10)

                if resultado[0]:
                    print(f"La ventana esta respondiendo (intento {intento}).")
                    return True
                else:
                    print(f"Intento {intento}: La ventana no responde aun. Esperando {intervalo}s...")
            except Exception:
                check = subprocess.run(
                    'tasklist /FI "IMAGENAME eq siep20.exe" /NH',
                    shell=True, capture_output=True, text=True
                )
                if "siep20.exe" not in check.stdout.lower():
                    print("SIEP ya no esta corriendo.")
                    return False
                print(f"Intento {intento}: No se pudo conectar a la ventana. Esperando {intervalo}s...")

            time.sleep(intervalo)

        print(f"TIMEOUT: La ventana no respondio en {timeout}s.")
        return False
        

    def presionar_boton_guardar(self, monto: float, cargo=False, valor_antes = 0.0):
        intentos = 0
        max_intentos = 6
        monto = round(monto, 2)
        valor_antes = round(valor_antes, 2)

        print(f"--- Iniciando proceso de guardado REAL (Monto objetivo: {monto}) ---")

        try:
            # Primer click en guardar antes de leer el valor
            print("Enviando primer comando de guardar...")
            pyautogui.click(388, 244)
            print("Esperando 10s a que el sistema procese...")
            time.sleep(10)

            while intentos < max_intentos:
                intentos += 1

                if cargo:
                    valor_actual = self.obtener_valor_campo_cargos()
                else:
                    valor_actual = self.obtener_valor_campo_abonos()

                if valor_actual is None:
                    print(f"Intento {intentos}: Lectura None. Esperando 15s respuesta del sistema...")
                    time.sleep(15)
                    continue

                valor_actual = round(valor_actual, 2)

                if valor_actual == valor_antes:
                    print(f"Intento {intentos}: Valor {valor_antes} igual al anterior. Enviando comando guardar...")
                    pyautogui.click(388, 244)

                    print("Esperando 10s a que el sistema procese el monto...")
                    time.sleep(10)
                    continue

                if not cargo:
                    if valor_actual > monto:
                        print(f"ERROR CRITICO: El abono {valor_actual} supera al monto {monto}")
                        return -2

                if valor_actual == monto:
                    print(f"EXCELENTE: Valor coincide ({valor_actual}). Voucher cuadrado.")
                    return True
                time.sleep(10)

            # Se agotaron los intentos: verificar si el programa sigue abierto
            if not self._ventana_sigue_abierta():
                print("ERROR: La ventana del voucher ya no existe. El programa se cerro.")
                return "VENTANA_CERRADA"

            print("Se agotaron los intentos de guardado sin exito.")
            return '-1'

        except Exception as e:
            print(f"Error en proceso real de guardado: {e}")
            return False

    def ingresar_registro(self, registro: str):
        try:
            v = self.voucher
            
            v.type_keys(registro, with_spaces=True)
            time.sleep(2)
            return True

        except Exception as e:
            print(f"Error de ingresar_registro: {e}")

    def guardar(self, p, test: bool = False, veces: int = 2):
        try:
            if test:
                return
            v = self.voucher
            for _ in range(veces):
                v.type_keys("{ENTER}")

            time.sleep(0.3)
            return True
        except Exception as e:
            print(f"[guardar]: {e}")

    def proceso_ot(self, orden_de: str, nro_documento: str):
        # Escribir el código de la OT
        for _ in range(2):
            pyautogui.press('tab')
            time.sleep(2)

        pyautogui.press('f1')
        dni, code = obtener_dni_cta_ot(nombre=orden_de)
        print(f"Codigo OT: {code}")
        escribir_tecla_por_tecla(texto=code)
        time.sleep(1)
        pyautogui.press('enter')
        time.sleep(0.1)
        pyautogui.press('enter')
        time.sleep(1)
        # Documento -> Otros
        pyautogui.press('tab')
        escribir_tecla_por_tecla(texto="Otros")
        time.sleep(1) 
        
        # Escribir el nro de documento
        for _ in range(2):
            pyautogui.press('tab')
            time.sleep(2)

        print(f"Nro Documento para escribir: {nro_documento}")
        escribir_tecla_por_tecla(texto=str(nro_documento))
        time.sleep(1) 
        
        # F1 y Nombre de la persona
        for _ in range(5):
            pyautogui.press('tab')
            time.sleep(1)
        
        pyautogui.press('f1')
        time.sleep(1)

        escribir_tecla_por_tecla(texto=dni)
        for _ in range(2):
            pyautogui.press('enter')
            time.sleep(0.1)
        
        pyautogui.press('tab')
        time.sleep(1)

    def unidad_operacion(self, ini: str):
        try:
            time.sleep(2.5)
            if ini.upper().startswith("M"):
                cantidad_down = 17 # TDA MEIGGS - VENTAS
            elif ini.upper().startswith("H"):
                cantidad_down = 13 # TDA HIDRAULICA VENTAS
            elif ini.upper().startswith("E"):
                cantidad_down = 12 # ELIAS AGUIRRE VENTAS
            
            v = self.voucher

            print("Celda objetiva en Unidad de unidad_operacion")

            if 0:
                print(f"Tipo de Ventas: {cantidad_down}")
                return

            for i in range(cantidad_down):
                v.type_keys("{DOWN}")
                time.sleep(0.2)

            return True
        except Exception as e:
            print(f"[unidad_operacion - error] {e}")

    def importe2(self, importe: str):
        try:
            v = self.voucher
            pyautogui.press('tab')
            time.sleep(1)
            pyautogui.press('tab')
            time.sleep(1)
            print(f"Importe para escribir: {importe}")
            v.type_keys(importe, with_spaces=True)
            time.sleep(1)
            v.type_keys("{DEL}{DEL}{DEL}")

            print("Importe ingresado")
            return True
        except Exception as e:
            print(f"[importe] {e}")

    def manipular_nieto_fast(self):
        try:
            print("=== Inicio Nieto ===")
            # Conexión directa a la ventana principal
            app = Application(backend="uia").connect(title_re=".*PRODUCTORES.*", timeout=10)
            abuelo = app.window(title_re=".*PRODUCTORES.*")
            
            print("Buscando ventana de impresión...")
            
            # Agregamos control_type para filtrar drásticamente la búsqueda
            nieto = abuelo.child_window(
                title_re=".*Impresión.*", 
                control_type="Window", # <--- CRUCIAL
                top_level_only=False
            )

            self.win_impresion = nieto.wait('exists', timeout=15)

            self.win_impresion.set_focus()
            print(f"✅ Conectado a: {self.win_impresion.window_text()}")
            return True

        except Exception as e:
            print(f"❌ No se encontró o tardó demasiado: {e}")
            return False

    def obtener_cantidad_filas(self):
        try:
            # Obtenemos todos los Edits de la ventana
            todos_los_edits = self.win_impresion.descendants(control_type="Edit")

            filas = 0
            fila_actual = {}

            for edit in todos_los_edits:
                nombre_interno = edit.window_text()

                # Detecta inicio de nueva fila
                if nombre_interno == "compr" and fila_actual:
                    filas += 1
                    fila_actual = {}

                fila_actual[nombre_interno] = edit

            if fila_actual:
                filas += 1

            time.sleep(2)
            return max(filas - 1, 0)
        except Exception as e:
            print(f"❌ No se pudo obtener filas: {e}")
            return -1

    def click_actualizar_tabla(self):
        try:
            # Buscamos por el ID interno con protección de duplicados
            btn_actualizar = self.win_impresion.child_window(
                auto_id="1016", 
                control_type="Button",
                found_index=0,
                top_level_only=False
            )
            
            btn_actualizar.set_focus()
            btn_actualizar.click_input()
            print("✅ Botón Actualizar (ID: 1016) presionado.")

            # Espera inteligente
            print("Esperando carga de datos...")
            for i in range(15):
                time.sleep(1)
                # Buscamos los Edits de forma segura
                edits = self.win_impresion.descendants(control_type="Edit")
                if edits and edits[0].window_text().strip() != "":
                    print(f"✅ Datos cargados en {i+1}s")
                    return True
            
            print("⚠️ Los datos no parecen haber cargado tras 15s")
            return False
        except Exception as e:
            print(f"❌ No se pudo clicar Actualizar: {e}")
            return False


    def click_modificar_tabla(self):
        try:
            app = Application(backend="uia").connect(title_re=".*PRODUCTORES.*", timeout=10)
            abuelo = app.window(title_re=".*PRODUCTORES.*")
            
            # Buscamos el botón Modificar (ID 1018)
            btn_modificar = abuelo.child_window(
                auto_id="1018", 
                control_type="Button",
                found_index=0,
                top_level_only=False
            )
            
            btn_modificar.set_focus()
            btn_modificar.click_input()
            print("✅ Click en Modificar (ID 1018)")
            
            time.sleep(5)
            
            pyautogui.click(141, 140)
            
            time.sleep(5)
            
            return True
        except Exception as e:
            print(f"[click_modificar_tabla] Error al presionar Modificar: {e}")
            return False


    def click_guardar_tabla(self):
        try:
            time.sleep(2)
            app = Application(backend="uia").connect(title_re=".*PRODUCTORES.*", timeout=10)
            abuelo = app.window(title_re=".*PRODUCTORES.*")
            
            btn_guardar = abuelo.child_window(
                auto_id="1019", 
                control_type="Button",
                found_index=0,
                top_level_only=False
            )
            
            btn_guardar.set_focus()
            btn_guardar.click_input()
            
            print("[click_guardar_tabla] Click en Guardar (ID 1019 - Index 0)")
            time.sleep(5)
            return True
            
        except Exception as e:
            print(f"❌ Error al presionar Guardar: {e}")
            return False

    def click_guardar_tabla_con_atajo(self, timeout_total=120, lote=None):
        try:
            print("Enfocado para guardar...?")

            resultado = [False]
            error = [None]

            def _todo():
                try:
                    app = Application(backend="uia").connect(title_re=".*PRODUCTORES.*", timeout=10)
                    app.window(title_re=".*PRODUCTORES.*").set_focus()
                    print("[click_guardar_tabla] Usando atajo Alt+G...")
                    pyautogui.hotkey('alt', 'g')
                    resultado[0] = True
                except Exception as e:
                    error[0] = e

            hilo = threading.Thread(target=_todo, daemon=True)
            hilo.start()
            hilo.join(timeout=timeout_total)

            if hilo.is_alive():
                print(f"TIMEOUT: click_guardar_tabla_con_atajo no respondio en {timeout_total}s. Programa congelado.")
                # Esperar a que se recupere
                if not self._esperar_que_responda(timeout=120, lote=lote):
                    return False
                return "RECUPERADO"

            if error[0]:
                print(f"Error durante guardar con atajo: {error[0]}")
                return False

            time.sleep(5)
            return True
        except Exception as e:
            print(f"[click_guardar_tabla_con_atajo] Error: {e}")
            return False


    def click_boton_agregar(self, timeout_click=120, lote: dict = None):
        try:
            time.sleep(2)
            app = Application(backend="uia").connect(title_re=".*PRODUCTORES.*", timeout=10)
            abuelo = app.window(title_re=".*PRODUCTORES.*")

            if not abuelo.is_active():
                abuelo.set_focus()

            btn_agregar = abuelo.child_window(
                auto_id="1007",
                control_type="Button",
                top_level_only=False
            )

            if not btn_agregar.exists():
                print("No se encontro el boton 'Agregar' con el ID 1007.")
                print("Esperando que el programa responda antes de intentar con coordenadas...")

                if not self._esperar_que_responda(timeout=120, lote=lote):
                    print("El programa no respondio. No se puede agregar.")
                    return False

                # El programa respondio, reintentar con pywinauto
                try:
                    app2 = Application(backend="uia").connect(title_re=".*PRODUCTORES.*", timeout=10)
                    abuelo2 = app2.window(title_re=".*PRODUCTORES.*")
                    btn2 = abuelo2.child_window(auto_id="1007", control_type="Button", top_level_only=False)
                    if btn2.exists():
                        btn2.click_input()
                        print("Boton 'Agregar' presionado despues de esperar.")
                        time.sleep(2)
                        return True
                except Exception:
                    pass

                # Si pywinauto sigue sin encontrarlo, intentar con coordenadas
                print("Intentando con coordenadas...")
                pyautogui.click(672, 147)
                print("Boton 'Agregar' presionado con coordenadas.")
                time.sleep(5)
                return "VERIFICAR"

            resultado = [False]
            error = [None]

            def _click():
                try:
                    btn_agregar.click_input()
                    resultado[0] = True
                except Exception as e:
                    error[0] = e

            hilo = threading.Thread(target=_click, daemon=True)
            hilo.start()
            hilo.join(timeout=timeout_click)

            if hilo.is_alive():
                print(f"TIMEOUT: click_boton_agregar no respondio en {timeout_click}s.")
                print("Esperando que el programa se recupere...")

                if not self._esperar_que_responda(timeout=120, lote=lote):
                    print("El programa no se recupero.")
                    ScreenImgEgresos.here(lote=lote, nombre="click_agregar_error_timeout")
                    time.sleep(5)
                    return False

                # Se recupero, intentar con coordenadas
                print("Intentando con coordenadas despues de recuperacion...")
                pyautogui.click(672, 147)
                time.sleep(5)
                return "VERIFICAR"

            if error[0]:
                print(f"Error durante click_input: {error[0]}")
                print("Esperando que el programa responda...")

                if not self._esperar_que_responda(timeout=120, lote=lote):
                    print("El programa no respondio despues del error.")
                    return False

                # Se recupero, reintentar con pywinauto
                try:
                    app2 = Application(backend="uia").connect(title_re=".*PRODUCTORES.*", timeout=10)
                    abuelo2 = app2.window(title_re=".*PRODUCTORES.*")
                    btn2 = abuelo2.child_window(auto_id="1007", control_type="Button", top_level_only=False)
                    if btn2.exists():
                        btn2.click_input()
                        print("Boton 'Agregar' presionado despues de recuperacion.")
                        time.sleep(4)
                        return True
                except Exception:
                    pass

                # Ultimo recurso: coordenadas
                print("Intentando con coordenadas...")
                pyautogui.click(672, 147)
                time.sleep(5)
                return "VERIFICAR"

            print("Boton 'Agregar' (ID: 1007) presionado correctamente.")
            time.sleep(3.5)
            return True

        except Exception as e:
            print(f"Error al intentar presionar 'Agregar': {e}")
            return False
    
    def agregar_datos_ultima_fila(self, cant: int = 20, lote: dict = None):
        intentos = 3
        cAdd = self.click_boton_agregar(lote=lote)
        if not cAdd:
            ScreenImgEgresos.here(lote=lote, nombre="click_agregar_error")
            time.sleep(5)
        UtilsModificar.reposionarse_a_ultima_fila(cant)
        
        for intento in range(1, intentos):
            print(f"Intento {intento}/{intentos} de agregar datos a la ultima fila...")
            
            celda = UtilsModificar.leer_celda()

            if celda:
                print("Celda con valor.")
                return True
            else:
                time.sleep(5)
                UtilsModificar.quitar_fila()

            #if intento == intentos:
                # 🔥 ÚLTIMO INTENTO
            UtilsModificar.click_cerrar_nieto()
            UtilsModificar.buscar_voucher()
            #UtilsModificar.click_guardar()
            UtilsModificar.click_modificar(1)

            UtilsModificar.nuevamente_click_en_cuenta()
            UtilsModificar.click_agregar_nuevamente()
            UtilsModificar.reposionarse_a_ultima_fila(cant)

        return False           

    def maximizar_ventana(self):
        try:
            self.win_impresion.maximize()
            print("✅ Ventana maximizada")
            time.sleep(4)
            return True
        except Exception as e:
            print(f"❌ No se pudo maximizar: {e}")

    def cerrar_ventana(self):
        try:
            self.win_impresion.close()
            print("✅ Ventana cerrada")
            time.sleep(2)
            return True
        except Exception as e:
            print(f"❌ Error al cerrar la ventana: {e}")

    def scroll_derecha_voucher_gui(self):
        try:
            x_derecha = 1011
            y_derecha = 711
            
            print(f"Ejecutando scroll con pyautogui en: {x_derecha}, {y_derecha}")
            
            pyautogui.PAUSE = 1 
            
            pyautogui.click(x=x_derecha, y=y_derecha, clicks=5, interval=1)
            
            print("✅ Scroll derecha completado con pyautogui")
            
        except Exception as e:
            print(f"❌ Error con pyautogui: {e}")

    def scroll_inicio_voucher_gui(self):
        try:
            x_izquierda = 10
            y_izquierda = 711
            
            print(f"Ejecutando scroll inicio con pyautogui en: {x_izquierda}, {y_izquierda}")
            
            pyautogui.click(x=x_izquierda, y=y_izquierda, clicks=8, interval=1)
            
            print("✅ Scroll devuelto al inicio (izquierda)")
            
        except Exception as e:
            print(f"❌ Error con pyautogui en scroll izquierdo: {e}")

    def scroll_derecha_voucher(self):
        try:
            # Aseguramos que la ventana esté al frente primero
            self.voucher.set_focus()
            
            # Coordenada Derecha: X=1011, Y=711
            # Usamos mouse.click para ignorar problemas de jerarquía UIA
            for _ in range(6):
                mouse.click(coords=(1011, 711))
                time.sleep(2.5)
                
            print("✅ Scroll derecha ejecutado en X=1011, Y=711")
        except Exception as e:
            print(f"❌ Error en scroll derecha: {e}")

    def click_columna_cliente(self, fila=1):
        try:
            print("Esperando 2 segundos antes de dar click en la nueva columna")
            time.sleep(2)
            edits = self.win_impresion.descendants(control_type="Edit")
            
            # 2. Buscar las celdas 'serie' y ordenarlas por posición vertical
            series = [e for e in edits if e.window_text() == "serie" and e.rectangle().top > 0]
            series.sort(key=lambda x: x.rectangle().top)

            if not series or fila > len(series):
                print(f"❌ No se encontró la columna 'serie' en la fila {fila}")
                return False

            print("Terminó la búsqueda?")

            # 3. Obtener el rectángulo de la serie en la fila N
            rect_s = series[fila - 1].rectangle()

            distancia_al_centro_cliente = 94
            
            x_final = rect_s.left - distancia_al_centro_cliente
            y_final = rect_s.mid_point().y

            # 5. Ejecución del clic físico
            self.win_impresion.set_focus()
            time.sleep(0.1) # Pausa para asegurar el foco
            
            pyautogui.click(x_final, y_final)
            print(f"✅ Clic en Cliente/Proveedor (Fila {fila}) realizado en: {x_final}, {y_final}")
            time.sleep(2)
            return True

        except Exception as e:
            print(f"❌ Error en click_columna_cliente: {e}")
            return False

    def escribir_columna_cliente(self, texto, fila=1, velocidad=0.5):
        try:
            if not self.click_columna_cliente(fila):
                return False
            
            time.sleep(0.2)

            for char in texto:
                pyperclip.copy(char)
                pyautogui.hotkey('ctrl', 'v')
                time.sleep(velocidad)
            
            pyautogui.press('down')
            
            print(f"✅ Texto '{texto}' escrito en Cliente/Proveedor (Fila {fila})")
            return True
            
        except Exception as e:
            print(f"❌ Error en escribir_columna_cliente: {e}")
            return False

    def click_campo_fila(self, nombre_columna="importe", numero_fila=1):
        print(f"\n--- BUSCANDO COLUMNA: '{nombre_columna}' | FILA: {numero_fila} ---\n")
        try:
            todos_los_edits = self.win_impresion.descendants(control_type="Edit")
            
            lista_elementos = []
            for edit in todos_los_edits:
                if edit.window_text() == nombre_columna:
                    rect = edit.rectangle()
                    if rect.top > 0:
                        lista_elementos.append((edit, rect.top))
            
            lista_elementos.sort(key=lambda x: x[1])
            
            total_encontrados = len(lista_elementos)
            print(f"Se encontraron {total_encontrados} campos con el nombre '{nombre_columna}'.")

            indice = numero_fila - 1
            if 0 <= indice < total_encontrados:
                target = lista_elementos[indice][0]
                
                print(f"✅ Fila {numero_fila} localizada. Haciendo clic...")
                self.win_impresion.set_focus()
                
                target.click_input()
                return True
            else:
                print(f"❌ La fila {numero_fila} no existe. (Máximo disponible: {total_encontrados})")
                return False

        except Exception as e:
            print(f"❌ Error al intentar clic dinámico: {e}")
            return False

    def escribir_campo_fila(self, valor=0.0, nombre_columna="importe", numero_fila=1):
        try:
            if not self.click_campo_fila(nombre_columna, numero_fila):
                return False
            
            time.sleep(0.2)
            pyautogui.hotkey('ctrl', 'a')
            time.sleep(0.1)

            texto = f"{valor:.2f}"  # → "1500.75", "200.00"
            pyperclip.copy(texto)
            pyautogui.hotkey('ctrl', 'v')
            
            pyautogui.press('down')
            
            print(f"✅ '{texto}' escrito en '{nombre_columna}' (Fila {numero_fila})")
            time.sleep(1)
            return True
            
        except Exception as e:
            print(f"❌ Error en escribir_campo_fila: {e}")
            return False

    def click_columna_cuenta(self, fila=1):
        try:
            # 1. Obtener la referencia de la celda Fecha (nuestra ancla)
            edits = self.win_impresion.descendants(control_type="Edit")
            fechas = [e for e in edits if e.window_text() == "fecha" and e.rectangle().top > 0]
            fechas.sort(key=lambda x: x.rectangle().top)

            if not fechas or fila > len(fechas):
                print(f"❌ No se encontró la fila {fila}")
                return False

            # 2. Obtener el rectángulo de la fecha en la fila deseada
            rect_f = fechas[fila - 1].rectangle()
            
            # 3. Cálculo de la coordenada exacta para 'Cuenta'
            # Iniciamos en el borde derecho de Fecha + la mitad del ancho de Cuenta (222 / 2)
            distancia_al_centro_cuenta = 111 
            
            x_final = rect_f.right + distancia_al_centro_cuenta
            y_final = rect_f.mid_point().y

            # 4. Ejecución del clic físico
            # Aseguramos que la ventana esté activa primero
            self.win_impresion.set_focus()
            time.sleep(0.1) # Breve pausa para evitar el doble clic accidental
            
            pyautogui.click(x_final, y_final)
            print(f"✅ Clic en Cuenta (Fila {fila}) realizado en: {x_final}, {y_final}")
            
            return True

        except Exception as e:
            print(f"❌ Error en click_columna_cuenta: {e}")
            return False

    def click_columna_cuenta_con_coordenadas(self, fila=1):
        try:
            x_final = 266
            y_final = 192
            pyautogui.click(x_final, y_final)
            print(f"✅ Clic en Cuenta (Fila {fila}) realizado en: {x_final}, {y_final}")
            time.sleep(1)
            return True

        except Exception as e:
            print(f"❌ Error en click_columna_cuenta: {e}")
            return False

    def obtener_tabla_mapeada(self):
        print("Mapeando tabla completa...")
        # Obtenemos todos los Edits de la ventana
        todos_los_edits = self.win_impresion.descendants(control_type="Edit")
        
        tabla_estructurada = []
        fila_actual = {}
        
        for i, edit in enumerate(todos_los_edits):
            nombre_interno = edit.window_text() # El nombre que sale en tu imagen
            
            # Si detectamos que se repite un nombre clave (ej. 'compr'), es una nueva fila
            if nombre_interno == "compr" and fila_actual:
                tabla_estructurada.append(fila_actual)
                fila_actual = {}
            
            fila_actual[nombre_interno] = edit
            
        # Añadimos la última fila
        if fila_actual:
            tabla_estructurada.append(fila_actual)
            
        print(f"✅ Se detectaron {len(tabla_estructurada)} filas en la tabla.")
        return tabla_estructurada


    def click_modificar_con_coordenadas(self):
        try:
            pyautogui.click(141, 142)
            time.sleep(15)
            pyautogui.click(141, 142)
        except Exception as e:
            print(f"[click_modificar_con_coordenadas] Error {e}")

    def click_modificar_tabla_con_atajo(self):
        try:
            # Nos aseguramos de que la ventana correcta esté activa (opcional pero recomendado)
            print("Enfocado para modificar...?")
            app = Application(backend="uia").connect(title_re=".*PRODUCTORES.*", timeout=10)
            main_window = app.window(title_re=".*PRODUCTORES.*")
            main_window.set_focus()
            
            # Hacemos clic en la ventana para asegurar el foco interno
            main_window.click_input(coords=(10, 10))
            time.sleep(0.5)

            print("[click_modificar_tabla] Usando atajo Alt+M...")
            pyautogui.hotkey('alt', 'm')
            time.sleep(5) # Mantenemos la espera por si la app tarda en reaccionar
            return True
        except Exception as e:
            print(f"[click_modificar_tabla_con_atajo] Error al enviar el atajo Alt+M: {e}")
            return False
        
    def cerrar_voucher(self):
        try:
            v = self.voucher
            v.close()
            print("Cierre correcto")
            time.sleep(2)
            return True
        except Exception as e:
            print(f"[cerrar_voucher] {e}")
    
    def finalizar(self):
        coordFinalizar = (441, 245)
        
        print(f"Haciendo click en FINALIZAR...")
        pyautogui.click(coordFinalizar)
        print(f"Click realizado en FINALIZAR")
        time.sleep(15)
        try:
            dlg = self.app.window(title="Guardando Voucher...")
            dlg.set_focus()
            print("Ventana 'Guardando Voucher' detectada. Confirmando cierre...")
            print("Los totales de cargos y abonos no cuadran?")
            time.sleep(8)
            pyautogui.press('s')
            time.sleep(15)
            return "CONTROL"
        except Exception as ee:
            print(f"ESTUVO BIEN LOS MONTOS?... {ee}")


    def imprimir(self, lote: dict = None):
        try:
            time.sleep(3)
            pyautogui.click(88, 145)
            time.sleep(3)
            pyautogui.click(611, 491)
            time.sleep(0.1)
            ScreenImgEgresos.here(lote=lote, nombre="IMPRESORA")
            time.sleep(4)
            return True
        except Exception as e:
            print(f"[imprimir] {e}")
            return False
