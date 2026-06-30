# -*- coding: utf-8 -*-
"""
Orquestador: descarga el crudo de Adryan y, si los datos CAMBIARON, corre el pipeline.

Flujo:
  1. Corre bot_adryan.py  -> descarga VACRptMotivo_*.xlsx a Descargas.
  2. Calcula un hash del CONTENIDO (ignorando la linea "Fecha y hora" del encabezado).
  3. Si el hash es distinto al de la ultima corrida -> corre pipeline.py.
     Si es igual -> NO corre el pipeline (evita reprocesar Excel sin necesidad).

Uso:
    python actualizar_todo.py            (normal: solo procesa si cambio)
    python actualizar_todo.py --forzar   (procesa siempre, aunque no cambie)
    python actualizar_todo.py --solo-bot (solo descarga, no procesa)
"""
import os
import sys
import json
import glob
import hashlib
import datetime
import subprocess
import warnings

warnings.filterwarnings("ignore")

AQUI = os.path.dirname(os.path.abspath(__file__))
PIPELINE_DIR = os.path.dirname(AQUI)                 # ...\PIPELINE
MOTOR = os.path.join(PIPELINE_DIR, "motor", "pipeline.py")
ESTADO = os.path.join(AQUI, "estado_bot.json")
LOG_DIR = os.path.join(AQUI, "logs")
os.makedirs(LOG_DIR, exist_ok=True)

PYEXE = sys.executable
CARPETA_DESCARGA = "C:/Users/jlopezp/Downloads"
PATRON = "VACRptMotivo_*.xlsx"


def log(msg: str):
    linea = f"[{datetime.datetime.now():%Y-%m-%d %H:%M:%S}] {msg}"
    print(linea, flush=True)
    with open(os.path.join(LOG_DIR, f"actualizar_{datetime.date.today():%Y%m%d}.log"),
              "a", encoding="utf-8") as f:
        f.write(linea + "\n")


def crudo_mas_reciente():
    archivos = glob.glob(os.path.join(CARPETA_DESCARGA, PATRON))
    return max(archivos, key=os.path.getmtime) if archivos else None


def hash_contenido(ruta: str) -> str:
    """Hash de las celdas, ignorando la fila 'Fecha y hora' (que cambia siempre)."""
    import openpyxl
    wb = openpyxl.load_workbook(ruta, read_only=True, data_only=True)
    ws = wb.active
    h = hashlib.sha256()
    for fila in ws.iter_rows(values_only=True):
        if fila and any(isinstance(c, str) and c.startswith("Fecha y hora") for c in fila):
            continue
        h.update(repr(fila).encode("utf-8", "ignore"))
    wb.close()
    return h.hexdigest()


def leer_estado() -> dict:
    if os.path.exists(ESTADO):
        try:
            return json.load(open(ESTADO, encoding="utf-8"))
        except Exception:
            return {}
    return {}


def guardar_estado(d: dict):
    json.dump(d, open(ESTADO, "w", encoding="utf-8"), ensure_ascii=False, indent=2)


def main():
    forzar = "--forzar" in sys.argv
    solo_bot = "--solo-bot" in sys.argv

    # 1) Descargar
    log("=== Paso 1: descargar crudo de Adryan ===")
    rc = subprocess.run([PYEXE, os.path.join(AQUI, "bot_adryan.py")]).returncode
    if rc != 0:
        log(f"El bot fallo (rc={rc}). Abortando.")
        return rc

    crudo = crudo_mas_reciente()
    if not crudo:
        log("No se encontro ningun crudo descargado. Abortando.")
        return 3
    log(f"Crudo mas reciente: {os.path.basename(crudo)}")

    if solo_bot:
        log("Modo --solo-bot: no se corre el pipeline.")
        return 0

    # 2) Detectar cambios
    nuevo_hash = hash_contenido(crudo)
    estado = leer_estado()
    if not forzar and estado.get("ultimo_hash") == nuevo_hash:
        log("Los datos NO cambiaron desde la ultima corrida -> NO se corre el pipeline.")
        estado["ultima_revision"] = datetime.datetime.now().isoformat(timespec="seconds")
        guardar_estado(estado)
        return 0

    # 3) Correr pipeline
    log("=== Paso 2: datos nuevos -> corriendo pipeline ===")
    rc = subprocess.run([PYEXE, MOTOR], cwd=os.path.dirname(MOTOR)).returncode
    if rc != 0:
        log(f"El pipeline termino con codigo {rc}.")
        return rc

    estado["ultimo_hash"] = nuevo_hash
    estado["ultimo_crudo"] = os.path.basename(crudo)
    estado["ultima_actualizacion"] = datetime.datetime.now().isoformat(timespec="seconds")
    guardar_estado(estado)
    log("=== LISTO: descarga + pipeline completados. ===")
    return 0


if __name__ == "__main__":
    sys.exit(main())
