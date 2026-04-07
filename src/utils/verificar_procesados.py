import json
import os
from src.config.path_folders import PathFolder


def todo_procesado() -> bool:
    
    path_json = os.path.join(PathFolder.FOLDER_PROCESO, 'lista.json')
    
    if not os.path.exists(path_json):
        print("[INFO] No existe el archivo JSON")
        return False

    with open(path_json, "r", encoding="utf-8") as f:
        data = json.load(f)

    total = len(data)
    no_proc = [i for i in data if i.get("procesado", "").lower() != "si"]

    if no_proc:
        print(f"[PENDIENTE] {len(no_proc)}/{total} registros NO procesados")
        return False

    print(f"[OK] {total}/{total} registros procesados")
    return True
