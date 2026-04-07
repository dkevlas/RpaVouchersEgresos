"""
Ejecutar como ADMINISTRADOR en el servidor:
    python exclusion_kaspersky.py
"""
import subprocess
import os
import sys

AVP = r"C:\Program Files (x86)\Kaspersky Lab\Kaspersky Small Office Security 21.24\avp.com"

if not os.path.exists(AVP):
    print(f"No se encontro avp.com en: {AVP}")
    print("Verifica la ruta de Kaspersky.")
    sys.exit(1)

# Rutas a excluir
exclusiones = [
    r"D:\SI_PROCASA\siep20.exe",
    r"D:\BOT_RPA_VOUCHER",
    sys.executable,  # python.exe actual
]

print("=" * 50)
print("AGREGANDO EXCLUSIONES EN KASPERSKY")
print("=" * 50)

for ruta in exclusiones:
    print(f"\nExcluyendo: {ruta}")
    cmd = f'"{AVP}" /C ExclusionTask /Add "{ruta}"'
    r = subprocess.run(cmd, shell=True, capture_output=True, text=True)
    print(f"  stdout: {r.stdout.strip()}" if r.stdout.strip() else "  stdout: (vacio)")
    print(f"  stderr: {r.stderr.strip()}" if r.stderr.strip() else "  stderr: (vacio)")
    print(f"  codigo: {r.returncode}")

# Intentar tambien con el parametro TRUSTED
print("\n" + "=" * 50)
print("INTENTANDO AGREGAR COMO APLICACION DE CONFIANZA")
print("=" * 50)

for ruta in exclusiones:
    if ruta.endswith(".exe") or ruta.endswith('"'):
        print(f"\nConfianza: {ruta}")
        cmd = f'"{AVP}" /C TrustedApps /Add "{ruta}"'
        r = subprocess.run(cmd, shell=True, capture_output=True, text=True)
        print(f"  stdout: {r.stdout.strip()}" if r.stdout.strip() else "  stdout: (vacio)")
        print(f"  stderr: {r.stderr.strip()}" if r.stderr.strip() else "  stderr: (vacio)")
        print(f"  codigo: {r.returncode}")

print("\n" + "=" * 50)
print("Si los comandos fallaron, intenta manualmente:")
print(f'  1. Abrir CMD como Administrador')
print(f'  2. Ejecutar: "{AVP}" HELP')
print(f"     Para ver los comandos disponibles")
print("=" * 50)
