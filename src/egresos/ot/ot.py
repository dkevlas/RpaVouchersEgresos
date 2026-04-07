import unicodedata
import json
import re
from typing import Dict, TypedDict, Optional, Tuple, List
from difflib import SequenceMatcher
from src.egresos.rutas.config import ConfigRutas

# ==========================
# BASE DE PERSONAS
# ==========================

class PersonaOT(TypedDict):
    DNI: str
    CTA: str


def obtener_personas_ot() -> List[Dict[str, PersonaOT]]:
    with open(ConfigRutas.OT_JSON, "r", encoding="utf-8") as file:
        return json.load(file)

PERSONAS_LISTA = obtener_personas_ot()

PERSONAS: Dict[str, PersonaOT] = {}

for item in PERSONAS_LISTA:
    PERSONAS.update(item)

def _normalizar(texto: str) -> str:

    texto = texto.upper()

    texto = unicodedata.normalize("NFD", texto)
    texto = texto.encode("ascii", "ignore").decode("utf-8")

    texto = re.sub(r"[^A-Z ]", "", texto)
    texto = re.sub(r"\s+", " ", texto)

    return texto.strip()


def _similitud(a: str, b: str) -> float:
    return SequenceMatcher(None, a, b).ratio()


PERSONAS_NORMALIZADAS = {
    _normalizar(nombre): codigo
    for nombre, codigo in PERSONAS.items()
}

def obtener_dni_cta_ot(nombre: str) -> Optional[Tuple[str, str]]:
    umbral = 0.90
    nombre_norm = _normalizar(nombre)

    mejor: Optional[PersonaOT] = None
    score_mejor = 0.0

    for base, datos in PERSONAS_NORMALIZADAS.items():
        score = _similitud(nombre_norm, base)

        if score > score_mejor:
            score_mejor = score
            mejor = datos

    if mejor and score_mejor >= umbral:
        return mejor["DNI"], mejor["CTA"]

    return None, None
