import os
import time
import pyautogui
from src.egresos.utils.verificar_fault import verificar_fault
from src.coordenadas.login import LoginCoordenadas
from src.schema.IEgresos import EgresoItem
from src.bot.steps_egresos import BotEgresos as Bot
from src.utils.cerrar_programa import cerrar_siep_seguro
from src.coordenadas.abrir_siep import abrir_siep_por_coordenadas
from src.egresos.utils.reintentar_desde_nieto import reintar_desde_nieto
from typing import List
from src.egresos.debug.leer_celda import UtilsModificar
from pynput.keyboard import Controller
from src.sistema.fault import cierre_fast
from src.egresos.excel.escritor_egresos import EscritorEgresos
from src.egresos.rutas.config import ConfigRutas
from src.egresos.datos.escribir_letra import escribir_tecla_por_tecla
from src.egresos.ot.mapeoEmpresas import obtener_clicks_por_empresa
from src.egresos.debug.screenshow import ScreenImgEgresos

keyboard = Controller()


def flujo_egresos(datos: List[EgresoItem]):

    cierre_fast()

    print(f"Cantidad de lotes a procesar: {len(datos)}")
    
    print("\n------ FLUJO EGRESOS ------\n")
    file_excel = [f for f in os.listdir(ConfigRutas.FOLDER_PROCESO) if f.lower().endswith('.xlsx')]
    if file_excel:
        inicial_file = os.path.join(ConfigRutas.FOLDER_PROCESO, file_excel[0])
    else:
        print(f'No se encontro archivo excel en {ConfigRutas.FOLDER_PROCESO}')
        return
    
    solo_file = os.path.basename(inicial_file)
    
    if solo_file.upper().startswith("MEE"):
        ini = 'MEE' # TDA MEIGGS - VENTAS
    elif solo_file.upper().startswith("HDE"):
        ini = 'HDE' # TDA HIDRAULICA VENTAS
    elif solo_file.upper().startswith("ELE"):
        ini = 'ELE' # ELIAS AGUIRRE VENTAS
    else:
        print("El archivo no comienza con el nombre esperado (MEE, HDE, ELE)")
        return "NO_FILE"

    print(f"INICIAL FILE: {solo_file}\n")

    MODO_MODIFICAR_TEST = False

    if True and not MODO_MODIFICAR_TEST:
        cerrar_siep_seguro()
        
        ok_open = abrir_siep_por_coordenadas()
        if ok_open == "NO_OPEN":
            time.sleep(3)
            return "NO_OPEN"
        
        if not ok_open:
            return "ERROR"
            
        LoginCoordenadas.escribir_usuario_contraseña_click()
        LoginCoordenadas.seleccionar_sucursal()
        
    AcumulativoError = 0
    try:
        bot = Bot()
        
        for i, lote in enumerate(datos, start=1):
            
            print("\n")
            print("=|"*15)
            print(f"[{i}/{len(datos)}] Procesando lote {lote.get('haber', {}).get('fecha_lote')}")
            print("=|"*15)
            print("\n")
            
            try:
                if lote.get('procesado') == "si":
                    print(f"SKIP: Lote {lote.get('haber', {}).get('fecha_lote')} ya procesado")
                    continue

                ScreenImgEgresos.mover_intentos_anteriores(lote)
                
                write_log = EscritorEgresos()
                
                print(f"\n\t===================== Comienza haber ======================\n")
                
                write_log.escribir_inicio(lote)
                
                if not MODO_MODIFICAR_TEST:
                    bot.openVoucherEgresos()
                    bot.focus_ventana_voucher()
                
                    ScreenImgEgresos.here(lote=lote, nombre="inicio", num=i, haber=True)
                    # ========================= HABER =========================
                    bot.llenar_cuenta(ini=ini)
                    bot.llenar_fecha(fecha=lote["haber"]["fecha"])
                    bot.actividad()
                    bot.flujo()
                    bot.seleccionar_documento()
                    bot.numero_documento(numero_documento=f"{ini}{lote['haber']['fecha_lote']}")
                    
                    isOt = lote["haber"]['tipo_doc'].upper() == 'OT'
                    
                    bot.ir_a_favor_de_control(cliente=lote["haber"]["orden_de"], ot=isOt)
        
                    time.sleep(5)
                    bot.ingresar_monto(monto=lote["haber"]["importe"])
                    bot.llenar_glosa(glosa=lote["haber"]["glosa"])

                    ScreenImgEgresos.here(lote=lote, nombre="guardar", num=i, haber=True)
                
                    time.sleep(5)
                    save_haber = bot.presionar_boton_guardar(monto=lote["haber"]["importe"])
                    print(f"Resultado guardar HABER: {save_haber} - [{type(save_haber)}]")
                    
                    if save_haber == "VENTANA_CERRADA":
                        write_log.escribir_observacion(lote, "El programa SIEP se cerro inesperadamente durante el guardado del HABER.")
                        ScreenImgEgresos.here(lote=lote, nombre=f"error_ventana_cerrada_haber", num=i, haber=True)
                        return "ERROR"

                    if save_haber == -2:
                        write_log.escribir_observacion(lote, "Error al registrar el monto en HABER: el valor es mayor al esperado, posiblemente por doble clic debido a la lentitud del sistema.")
                        print(f"Error al guardar monto HABER: {lote['haber']['importe']}")
                        ScreenImgEgresos.here(lote=lote, nombre=f"error_guardar_haber_monto_superior", num=i, haber=True)
                        return "ERROR"

                    if save_haber == '-1':
                        print(f"Error al guardar monto HABER: {lote['haber']['importe']}")
                        write_log.escribir_observacion(lote, "Error al registrar el monto en HABER: se agotaron los intentos de guardado sin éxito.")
                        ScreenImgEgresos.here(lote=lote, nombre=f"error_guardar_haber_intentos_fallidos", num=i, haber=True)
                        return "ERROR"
                    
                    IMPORTE_CARGO = 0
                    
                    # ========================= DEBER =========================
                    print(f"\n\t================ Comienza deber ======================\n")
                    
                    total_deber = len(lote.get('deber', []))
                    for i, deber in enumerate(lote.get('deber', []), start=1):

                        print(f"\n\t[{i}/{len(lote.get('deber', []))}]\n")
                        
                        if deber.get('tipo_doc', '').upper() == 'OT':

                            bot.proceso_ot(
                                orden_de=deber['orden_de'],
                                nro_documento=deber['nro_documento']
                            )
                            bot.unidad_operacion(ini=ini)
                            bot.importe2(importe=deber['importe'])
                            
                        else:
                            bot.ingresar_registro(registro=deber['registro'])
                            for o in range(11):
                                pyautogui.press('tab')
                                if o == 0:
                                    time.sleep(1) 
                                time.sleep(0.5)
                            bot.unidad_operacion(ini=ini)
                            bot.importe2(importe=deber['importe'])

                        print(f"\nImporte antes: {IMPORTE_CARGO}")
                        
                        importe_antes = IMPORTE_CARGO
                        IMPORTE_CARGO += deber['importe']
                        
                        print(f"Importe despues: {IMPORTE_CARGO}\n")
                        
                        save_monto = bot.presionar_boton_guardar(monto=IMPORTE_CARGO, cargo=True, valor_antes=importe_antes)
                        if save_monto == "VENTANA_CERRADA":
                            write_log.escribir_observacion(lote, "El programa SIEP se cerro inesperadamente durante el guardado del DEBER.")
                            ScreenImgEgresos.here(lote=lote, nombre=f"error_ventana_cerrada_deber", num=i, deber=True)
                            return "ERROR"

                        if save_monto == -2:
                            write_log.escribir_observacion(lote, "Error al registrar el monto en DEBER: el valor es mayor al esperado, posiblemente por doble clic debido a la lentitud del sistema.")
                            print(f"Error al guardar monto {IMPORTE_CARGO}")
                            ScreenImgEgresos.here(lote=lote, nombre=f"error_guardar_deber_monto_superior", num=i, haber=False)
                            return "ERROR"

                        if save_monto == '-1':
                            print(f"Error al guardar monto {IMPORTE_CARGO}")
                            write_log.escribir_observacion(lote, "Error al registrar el monto en DEBER: se agotaron los intentos de guardado sin éxito.")
                            ScreenImgEgresos.here(lote=lote, nombre=f"error_guardar_deber_intentos_fallidos", num=i, haber=False)
                            return "ERROR"

                        if i < total_deber:
                            time.sleep(2)

                    time.sleep(3)
                    
                    fin = bot.finalizar()
                    
                    ScreenImgEgresos.here(lote=lote, nombre=f"finalizar", num=i, haber=False)
                    
                    print(f"\nResultado finalizar: {fin}")
                    
                    tiempo = 25
                    tiempo_extra = 0 if len(lote.get('deber', [])) > 15 else 0
                    tiempo_total = tiempo + tiempo_extra
                    print(f"\n\tTiempo de espera dado: {tiempo_total}\n")
                    time.sleep(tiempo_total)

                result = bot.manipular_nieto_fast()

                print(f"\nRESULTADO FINAL VENTANA NIETA: {result}")

                opcion2 = False
                if not result:
                    print("No se detectó ventana nieta")
                    ScreenImgEgresos.here(lote=lote, nombre=f"ventana_nieta_no_detectada", num=i)
                    print("Cerrando ventana nieta")
                    pyautogui.click(900, 213)
                    time.sleep(5)
                    ScreenImgEgresos.here(lote=lote, nombre=f"ventana_nieta_cerrada", num=i)
                    
                    UtilsModificar.buscar_voucher()

                    result = True
                    opcion2 = True

                ScreenImgEgresos.here(lote=lote, nombre=f"impresion_voucher", num=i)
                
                if not opcion2:
                    print("Simulacion Abrir ventana nieta")
                    bot.maximizar_ventana()

                if result:
                    if lote.get('agregar', []):

                        if not opcion2:
                            bot.click_modificar_tabla()
                        else:
                            UtilsModificar.click_modificar()
                    
                        ScreenImgEgresos.here(lote=lote, nombre=f"modificar_tabla", num=i)
                
                        num_filas_actual = bot.obtener_cantidad_filas()
                        print(f"TABLAS numero de FILAS: {num_filas_actual}")
                
                    if not opcion2:
                        bot.maximizar_ventana()
                        print("Se maximizó ventana")
                    
                    # ================== AGREGAR ==================

                    print("\n ====================  COMIENZO AGREGAR ================ \n")
                    
                    for add_index, add in enumerate(lote.get('agregar', []), start=1):
                        time.sleep(3)
                        print(f"\n\t[{add_index}/{len(lote.get('agregar', []))}]\n")

                        bot.click_columna_cuenta_con_coordenadas()

                        print("Click segunda columna, primera fila")
                        cant_deber = len(lote.get('deber', []))
                        cant_agregar = len(lote.get('agregar', []))

                        cant_total = cant_deber + cant_agregar + 1
                        print(f"CANT TOTAL: {cant_deber} + {cant_agregar} + 1 = {cant_total}")
                        
                        add_fila = bot.agregar_datos_ultima_fila(cant_total)

                        if not add_fila:
                            print("Error al agregar nueva fila en Modificar")
                            ScreenImgEgresos.here(lote=lote, nombre=f"error_agregar_modificar")
                            write_log.escribir_observacion(lote, "Error al agregar nueva fila en Modificar")
                            time.sleep(3)
                            return "ERROR"

                        for _ in range(3):
                            pyautogui.press('tab')
                            time.sleep(1)

                        escribir_tecla_por_tecla(add["orden_de"].lower()[:33])
                        num_down = obtener_clicks_por_empresa(add["orden_de"])
                        time.sleep(1)
                        for _ in range(num_down):
                            pyautogui.press('down')
                            time.sleep(0.5)

                        time.sleep(2)
                        for _ in range(6):
                            pyautogui.press('tab')
                            time.sleep(1.5)
                            
                        print("Simulación de escribir moneda soles....")
                        pyautogui.write(f"{add['importe']}", interval=0.3)

                        for _ in range(1):
                            pyautogui.press('tab')
                            time.sleep(1.5)
                        
                        print("Simulación de escribir moneda dolares....")
                        pyautogui.write(f"{0.0}", interval=0.3)

                        time.sleep(2)

                        print("Scroll inicio voucher gui")
                        bot.scroll_inicio_voucher_gui()
                        
                        if verificar_fault(lote, write_log) == "ERROR": return "ERROR"
                        
                        time.sleep(2)

                        print("Click guardar tabla con atajo")
                        bot.click_guardar_tabla_con_atajo(lote=lote)
                        
                        time.sleep(5)
                        
                        if add_index < len(lote.get('agregar', [])):
                            print("Click modificar tabla con atajo")
                            bot.click_modificar_con_coordenadas()

                            time.sleep(5)

                    # ================== FIN AGREGAR ==================

                    bot.click_guardar_tabla_con_atajo()
                    
                    ScreenImgEgresos.here(lote=lote, nombre=f"MODIFICAR", add=True)
                else:
                    print("ERROR: No se pudo enfocar la ventana Impresión de Vouchers")
                    AcumulativoError += 1
                    write_log.escribir_observacion(lote, f"ERROR: No se pudo finalizar el voucher")
                    return -1

                ScreenImgEgresos.here(lote=lote, nombre=f"FIN")

                if not MODO_MODIFICAR_TEST:
                    print("Preparando para imprimir")
                    bot.imprimir(lote=lote)

                if not bot.cerrar_ventana():
                    print("Error al cerrar ventana nieta")
                    ScreenImgEgresos.here(lote=lote, nombre=f"error_Cerrar_ventana_nieta")
                    return "ERROR"

                bot.cerrar_voucher()

                intento_actual = lote.get("intento", 1)
                obs_reintento = f"Procesado correctamente en intento {intento_actual}" if intento_actual > 1 else ""
                write_log.escribir_fin(lote, observacion=obs_reintento)

                ScreenImgEgresos.here(lote=lote, nombre=f"cerrar_todo")
                
                time.sleep(8)

            except Exception as e:
                print(f"[ERROR LOTE] {e}")
                input("PAUSA STOP BYE")
                resultado = verificar_fault(lote, write_log)
                if resultado == "ERROR": return "ERROR"
                # Si verificar_fault no detectó error en SIEP, registrar la excepción como observación
                write_log.escribir_observacion(lote, f"Error inesperado: {e}")
                ScreenImgEgresos.here(lote=lote, nombre=f"error_inesperado", num=i)
                return "ERROR"

        else:
            print("Todo procesado OK")            
            return "OK"

    except Exception as e:
        print(f"ERROR_ FLUJOS EGRESOS: {e}")     
        return "ERROR_FLUJO"
