import json
from difflib import SequenceMatcher
from src.egresos.rutas.config import ConfigRutas
def similitud(a: str, b: str) -> float:
    return SequenceMatcher(None, a.lower(), b.lower()).ratio()

def obtener_clicks_por_empresa(orden_de: str):
    UMBRAL = 0.90
    if not orden_de:
        return None

    with open(ConfigRutas.MAPEO_EMPRESAS, "r", encoding="utf-8") as f:
        data = json.load(f)

    mejor_match = None
    mejor_score = 0.0

    for emp in data["empresas"]:
        score = similitud(orden_de, emp["nombre"])
        if score > mejor_score:
            mejor_score = score
            mejor_match = emp

    if mejor_match and mejor_score >= UMBRAL:
        return mejor_match["clicks"]

    return 1
