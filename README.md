# Bot RPA - Vouchers de Egresos (SIEP)

Bot de automatizacion robotica de procesos (RPA) que registra vouchers de egresos en el sistema **SIEP** (Sistema Integrado de Empresas Publicas) a partir de archivos Excel.

## Descripcion

El bot lee archivos Excel con datos de vouchers financieros, los convierte a JSON y automatiza el llenado de formularios en la aplicacion de escritorio SIEP mediante interaccion directa con la interfaz grafica de Windows.

### Flujo general

1. Verifica que no exista otra instancia en ejecucion (lock file)
2. Lee los archivos Excel desde la carpeta de datos preliminares
3. Convierte los datos a JSON para procesamiento interno
4. Abre SIEP e inicia sesion automaticamente
5. Completa los formularios de vouchers (cuenta, fecha, documento, monto, glosa)
6. Envia notificaciones por correo en cada etapa del proceso
7. Mueve los archivos procesados a la carpeta correspondiente

## Estructura del proyecto

```nginx
├── app.py                          # Punto de entrada principal
├── bot_egresos.py                  # Orquestador del bot (reintentos, lock, correos)
└── src/
    ├── flujo_egresos.py            # Flujo principal de procesamiento
    ├── bot/
    │   └── steps_egresos.py        # Automatizacion de ventanas SIEP (pywinauto)
    ├── config/
    │   ├── path_folders.py         # Configuracion de rutas de carpetas
    │   └── get_ruta_base.py        # Ruta base (ejecutable vs desarrollo)
    ├── coordenadas/
    │   ├── login.py                # Login automatico en SIEP
    │   └── abrir_siep.py           # Apertura de la aplicacion SIEP
    ├── egresos/
    │   ├── datos/                  # Obtencion y preparacion de datos
    │   ├── excel/                  # Conversion Excel <-> JSON
    │   ├── ot/                     # Mapeo de ordenes y empresas
    │   ├── mensaje/                # Plantillas HTML de correo
    │   ├── utils/                  # Utilidades de egresos
    │   └── debug/                  # Capturas de pantalla y diagnostico
    ├── excel/                      # Procesamiento Excel (ingresos)
    ├── mensaje/
    │   └── enviar_mensaje.py       # Envio de correo SMTP + fallback Outlook
    ├── schema/                     # Definiciones de tipos (TypedDict)
    ├── sistema/                    # Manejo de eventos y cierre del sistema
    ├── log/                        # Configuracion de logging
    ├── utils/                      # Utilidades generales
    └── test/                       # Scripts de prueba
```

## Requisitos

- **Python 3.8+**
- **Windows** (requiere interaccion con GUI de escritorio)
- **SIEP** instalado en el equipo

### Dependencias

```
pyautogui
pywinauto
pynput
pyperclip
openpyxl
pywin32
```

## Estructura de carpetas de datos

```
D:\BOT_RPA_VOUCHER\Egresos\
├── Data Preliminar/    # Archivos Excel de entrada
├── Proceso/            # Archivos en procesamiento (JSON + Excel)
└── Data Procesada/     # Archivos completados
```

### Nomenclatura de archivos

Los archivos Excel deben comenzar con un prefijo que indica la tienda:

| Prefijo | Tienda |
|---------|--------|
| `MEE`   | TDA MEIGGS - VENTAS |
| `HDE`   | TDA HIDRAULICA VENTAS |
| `ELE`   | ELIAS AGUIRRE VENTAS |

## Uso

```bash
python app.py
```

El bot procesara hasta **3 registros** por ejecucion (configurable en `app.py`). Incluye:

- **Lock file** para evitar ejecuciones simultaneas (expira a las 4 horas)
- **2 reintentos** automaticos en caso de fallo
- **Sincronizacion memoria/disco** entre reintentos para evitar duplicados
- **Notificaciones por correo** al inicio, fin y en caso de error

## Notificaciones

El bot envia correos electronicos via SMTP en los siguientes eventos:

- Inicio de ejecucion (con detalle de registros a procesar)
- Proceso finalizado exitosamente
- Proceso detenido o con errores
- Imposibilidad de abrir SIEP

## Notas

- El bot utiliza coordenadas de pantalla y automatizacion de ventanas, por lo que la resolucion de pantalla y la posicion de las ventanas deben mantenerse consistentes.
- No ejecutar otras aplicaciones que interfieran con el foco de la ventana de SIEP durante el procesamiento.
- Las credenciales de acceso estan configuradas internamente en el codigo.
