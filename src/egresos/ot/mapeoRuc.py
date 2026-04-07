import json
import re
from difflib import SequenceMatcher
from src.egresos.rutas.config import ConfigRutas


def similitud(a: str, b: str) -> float:
    return SequenceMatcher(None, a.lower(), b.lower()).ratio()


def tiene_caracteres_especiales(texto: str) -> bool:
    return bool(re.search(r"[^a-zA-Z0-9\s.,\-]", texto))


def obtener_ruc_si_especial(cliente: str):
    """
    Si el nombre del cliente tiene caracteres especiales (como '),
    busca en MapeoRuc.json y retorna el RUC para escribirlo en su lugar.
    Si no tiene caracteres especiales, retorna None (se usa el nombre normal).
    """
    if not tiene_caracteres_especiales(cliente):
        return None

    UMBRAL = 0.85

    try:
        with open(ConfigRutas.MAPEO_RUC, "r", encoding="utf-8") as f:
            data = json.load(f)
    except FileNotFoundError:
        print(f"[mapeoRuc] No se encontro {ConfigRutas.MAPEO_RUC}")
        return None

    mejor_match = None
    mejor_score = 0.0

    for emp in data.get("empresas", []):
        score = similitud(cliente, emp["nombre"])
        if score > mejor_score:
            mejor_score = score
            mejor_match = emp

    if mejor_match and mejor_score >= UMBRAL:
        ruc = mejor_match.get("ruc", "")
        if ruc:
            print(f"[mapeoRuc] '{cliente}' -> RUC: {ruc} (similitud: {mejor_score:.2f})")
            return ruc
        else:
            print(f"[mapeoRuc] '{cliente}' encontrado pero sin RUC asignado en MapeoRuc.json")
            return None

    print(f"[mapeoRuc] '{cliente}' tiene caracteres especiales pero no se encontro en MapeoRuc.json")
    return None
