from typing import TypedDict, List, Optional


class DeberItem(TypedDict):
    fecha: str
    fecha_lote: str
    tipo_doc: str
    nro_documento: str
    orden_de: str
    registro: str
    importe: float
    glosa: str
    fila: int


class HaberItem(TypedDict):
    fecha: str
    fecha_lote: str
    tipo_doc: str
    nro_documento: str
    orden_de: str
    importe: float
    glosa: str
    fila: int


class EgresoItem(TypedDict):
    procesado: str
    inicio: str
    fin: str
    observacion: str
    haber: HaberItem
    deber: List[DeberItem]
    agregar: List[DeberItem]


class IReturnEgresos(TypedDict, total=False):
    success: bool
    message: str
    data_json: Optional[List[EgresoItem]]
    stop: Optional[bool]
    message_outlook: Optional[bool]
    asunto: Optional[str]
