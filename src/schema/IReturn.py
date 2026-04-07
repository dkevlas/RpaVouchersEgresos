from typing import TypedDict, Optional


class RegistroItem(TypedDict):
    Fecha: str
    Serie: str
    Nro: str
    Cliente: str
    Tipo_Pago: str
    Total: float
    Monto: float
    nro_operacion: str
    posicion: int
    procesada: str
    duracion: int
    inicio: str
    final: str
    observacion: str
    

class IReturn(TypedDict):
    success: bool
    message: str
    data_json: Optional[RegistroItem] = []
    stop: Optional[bool] = False
    message_outlook: Optional[bool] = False
