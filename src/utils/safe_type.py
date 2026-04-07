import time


def safe_type(dlg, keys, timeout=120, pause=0.05, with_spaces=False):
    start = time.time()
    while True:
        try:
            dlg.type_keys(keys, pause=pause, with_spaces=with_spaces)
            return True
        except Exception as e:
            if time.time() - start > timeout:
                print(f"[safe_type] Timeout enviando {keys}")
                return False
            print(f"[safe_type] La app no responde, reintentando... {keys}")
            time.sleep(1)
