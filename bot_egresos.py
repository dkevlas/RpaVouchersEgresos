import os
import time
import json
import traceback
from src.config.path_folders import PathFolder
from src.utils.cerrar_programa import cerrar_siep_seguro
from src.egresos.utils.mover_si_todo_procesado import mover_si_todo_procesado_egresos
from src.flujo_egresos import flujo_egresos as flujo
from src.egresos.mensaje.msg_no_continua import crear_msg_egresos_no_continua
from src.egresos.mensaje.msg_finalizado import crear_msg_egresos_finalizado
from src.egresos.mensaje.msg_inicio import crear_msg_egresos_inicio
from src.mensaje.enviar_mensaje import enviar_correo_smtp
from src.egresos.mensaje.msg_no_continua import crear_msg_egresos_no_continua
from src.sistema.eventos import mensaje_eventos
from typing import List
from src.egresos.datos.obtener_datos_egresos import obtener_datos_egresos
from src.utils.eliminar_log import eliminar_log
from src.sistema.fault import cierre_fast

LOCK_FILE = os.path.join(PathFolder.FOLDER_BASE, 'app.lock')

ELIMINAR = True

def botVouchersEgresos(cant: int = 40, path_log: str = None):
    global ELIMINAR
    try:        
        time.sleep(2)

        MAX_HORAS = 4
        MAX_SEGUNDOS = MAX_HORAS * 3600

        if os.path.exists(LOCK_FILE):
            antiguedad = time.time() - os.path.getmtime(LOCK_FILE)

            if antiguedad > MAX_SEGUNDOS:
                print("Lock antiguo (>4h). Eliminando...")
                os.remove(LOCK_FILE)
            else:
                print(f'Ya existe una ejecución... Deteniendo flujo')
                ELIMINAR = False
                eliminar_log(path_log)
                time.sleep(5)
                os._exit(0)

        with open(LOCK_FILE, "w") as f:
            print("Escribiendo lock desde BotEgresos...")
            f.write("lock")

        mensaje_eventos()
        
        results = obtener_datos_egresos(n=cant) # CANTIDAD DE DATOS PARA PROCESAR 

        if not results.get('success'):
            print(results.get('message'))
            
            os.remove(LOCK_FILE)

            if results.get('message_outlook'):
                enviar_correo_smtp(
                    body=crear_msg_egresos_no_continua(motivo=results.get('message')),
                    asunto=results.get("asunto")
                )
            
            eliminar_log(path_log)
            time.sleep(15)
            os._exit(0)    
        
        datos: List[dict] = results.get('data_json', [])

        enviar_correo_smtp(
            body=crear_msg_egresos_inicio(datos=datos),
            asunto='Egresos: Inicio de ejecución del bot'
        )

        print(f'\nSe procesarán {len(datos or [])} registros\n')

        intentos = 2
        for i in range(intentos):
            print(f'\nIntento {i+1}/{intentos} del flujo del bot\n')
            
            if i > 0:
                print("Sincronizando memoria con disco para evitar duplicados...")
                try:
                    path_json_sync = os.path.join(PathFolder.FOLDER_PROCESO, 'egresos.json')
                    if os.path.exists(path_json_sync):
                        with open(path_json_sync, 'r', encoding='utf-8') as f:
                            data_disk = json.load(f)

                        # Mapa clave -> procesado (por tipo_doc + nro_documento del haber)
                        mapa_disk = {
                            (str(d.get('haber', {}).get('tipo_doc')), str(d.get('haber', {}).get('nro_documento'))): d.get('procesado')
                            for d in data_disk
                        }

                        for d_mem in datos:
                            haber = d_mem.get('haber', {})
                            k = (str(haber.get('tipo_doc')), str(haber.get('nro_documento')))
                            if k in mapa_disk:
                                d_mem['procesado'] = mapa_disk[k]
                        print("Sincronización completada.")
                except Exception as e:
                    print(f"Advertencia al sincronizar memoria: {e}")

            print("\n")
            
            cierre_fast()

            for l in datos:
                print(f"FECHA: {l['haber']['fecha']}")

            print("\n")

            result = flujo(datos=datos)
            
            print(f'\nResultado intento {i+1}/{intentos}: {result}\n')

            if result == "OK": break
            
            if result == "NO_OPEN": break

            if result == "ERROR":
                cierre_fast()
            else:
                cierre_fast()

            if result == "NO_FILE":
                break
            
            print(f'\nIntento {i+1}/{intentos} del flujo del bot fallido\n')

        # Mensaje final (unico)
        try:
            if result == "OK":
                enviar_correo_smtp(
                    body=crear_msg_egresos_finalizado(datos=datos),
                    asunto='Egresos: Proceso finalizado'
                )
            elif result == "NO_OPEN":
                enviar_correo_smtp(
                    body=crear_msg_egresos_no_continua(motivo="No se pudo abrir SIEP despues de varios intentos."),
                    asunto='Egresos: No se pudo abrir SIEP'
                )
            elif result == "NO_FILE":
                enviar_correo_smtp(
                    body="",
                    asunto="Egresos: No se encontró archivo con nombre inicial correcto"
                )
            else:
                enviar_correo_smtp(
                    body=crear_msg_egresos_finalizado(datos=datos),
                    asunto='Egresos: Proceso finalizado con observaciones'
                )
        except Exception as e:
            print(f"Error al enviar correo final: {e}")


        time.sleep(1)

        mover_si_todo_procesado_egresos()

        time.sleep(10)
    except KeyboardInterrupt as k:
        print(f'\n\n[TECLADO] interrupcion {k}\n')
        time.sleep(10)
    except Exception as e:
        print(f'\n\n[ERROR GLOBAL]\n\n{e}\n\n')
        traceback.print_exc()
        time.sleep(10)
    finally:
        try:
            if ELIMINAR:
                print("Eliminando lock desde BotEgresos...")
                os.remove(LOCK_FILE)
                cerrar_siep_seguro()
        except Exception as e:
            print(f"Error al cerrar los logs o el app.lock desde BotEgresos?: {e}")
