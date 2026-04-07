import json
import shutil
import os
from typing import List
from src.egresos.rutas.config import ConfigRutas
from src.egresos.excel.actualizar_json_egresos_por_excel import actualizar_json_por_excel
from src.schema.IEgresos import IReturnEgresos, EgresoItem
from src.egresos.excel.xls_a_xlsx import xls_a_xlsx
from src.egresos.excel.crear_json import crear_json_egresos


def get_dato(filejson: str) -> List[EgresoItem]:
    with open(filejson, "r", encoding="utf-8") as f:
        data: List[EgresoItem] = json.load(f)

    return [i for i in data if i.get("procesado") == "no"]


def obtener_datos_egresos(n: int) -> IReturnEgresos:
    try:
        filename_json = os.path.join(ConfigRutas.FOLDER_PROCESO, 'egresos.json')
        
        if os.path.exists(filename_json):
            print(f"Se encontró un json en {ConfigRutas.FOLDER_PROCESO}")
            files_excel = [e for e in os.listdir(ConfigRutas.FOLDER_PROCESO) if e.lower().endswith('.xlsx')]
            if files_excel:
                actualizar_json_por_excel()
                msg = 'Existe Json y se actualizó'
                print(msg)
                return {'success': True, 'message': msg, 'data_json': get_dato(filejson=filename_json)[:n]}
            else:
                msg = 'Se continuará con el json existente'
                return {'success': True, 'message': msg, 'data_json': get_dato(filejson=filename_json)[:n]}
        
        else:
            print(f"No se encontró un json en {ConfigRutas.FOLDER_PROCESO}")
            print(f"Se buscará un excel en {ConfigRutas.FOLDER_PRELIMINAR}")
            
            files_excel = [e for e in os.listdir(ConfigRutas.FOLDER_PRELIMINAR) if e.lower().endswith(('.xlsx', 'xls'))]
            if not files_excel:
                msg = f'\nNo existe un excel en {ConfigRutas.FOLDER_PRELIMINAR}'
                return {'success': False, 'message': msg, 'stop': True}
            else:

                filename_excel = os.path.basename(files_excel[0])
                if filename_excel[:3].upper() not in ("MEE", "HDE", "ELE"):
                    msg = "El archivo no empieza con MEE, HDE o ELE"
                    print(msg)
                    return {
                        "success": False,
                        "message": msg,
                        "data_json": [],
                        "message_outlook": True,
                        "stop": True,
                        "asunto": "Egresos: No se encontró archivo con nombre inicial correcto"
                    }

                dir_proceso = ConfigRutas.FOLDER_PROCESO
                path_excel = os.path.join(ConfigRutas.FOLDER_PRELIMINAR, files_excel[0])

                path_excel = xls_a_xlsx(ruta_xls=path_excel)

                print(f'\nMoviendo {path_excel} -> {dir_proceso}')
                
                shutil.move(path_excel, dir_proceso)
                new_ruta_excel = os.path.join(dir_proceso, [e for e in os.listdir(dir_proceso) if e.lower().endswith('.xlsx')][0])
                print(f'Nueva ruta de excel: {new_ruta_excel}')
                result = crear_json_egresos(ruta_excel=new_ruta_excel)
                
                print("="*50)
                print(result)
                print("="*50)

                file_json = [j for j in os.listdir(ConfigRutas.FOLDER_PROCESO) if j.lower().endswith('.json')]
                if file_json:
                    return {
                        'success': True,
                        'message': 'Se creó un nuevo json',
                        'data_json': get_dato(filejson=os.path.join(dir_proceso, file_json[0]))[:n],
                        'message_outlook': False,
                        'stop': False
                    }
                else:
                    return {
                        'success': False,
                        'message_outlook': True,
                        'stop': True,
                        'message': 'Error desconocido al crear un json nuevo'
                    }
    except Exception as e:
        msg = f'Error desconocido en obtener los datos para procesar: {str(e)}'
        return {
            'message': msg,
            'data_json': None,
            'message_outlook': True,
            'stop': True,
            'success': False,
            "asunto": msg
        }
