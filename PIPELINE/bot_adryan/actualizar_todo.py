# -*- coding: utf-8 -*-
"""
Orquestador diario: descarga el crudo de Adryan (vacaciones) y el Maestro de
Personal, y corre el pipeline si los datos de vacaciones CAMBIARON.

Flujo:
  1. Corre bot_adryan.py  -> descarga VACRptMotivo_*.xlsx a Descargas.
  2. Calcula un hash del CONTENIDO (ignorando la linea "Fecha y hora" del encabezado).
  3. Si el hash es distinto al de la ultima corrida -> corre pipeline.py.
     Si es igual -> NO corre el pipeline (evita reprocesar Excel sin necesidad).
  4. Corre bot_maestro.py -> descarga y sanitiza el Maestro de Personal (DATAS/).
     Siempre corre (no tiene el mismo chequeo de "cambio" que vacaciones: es
     mas barato y el maestro se usa para filtrar fantasmas, conviene tenerlo
     fresco siempre que se pueda).

Pensado para dispararse una vez al dia desde servidor.py al arrancar (ver
_disparar_actualizacion_diaria_si_corresponde en servidor.py), no por un
horario fijo de Windows: si la PC estaba apagada a esa hora, la tarea
programada nunca hubiera corrido. Aqui en cambio se evalua "¿ya corrio hoy?"
en cada arranque de la app, sin importar la hora.

Uso:
    python actualizar_todo.py            (normal: solo procesa vacaciones si cambio; maestro siempre)
    python actualizar_todo.py --forzar   (procesa vacaciones siempre, aunque no cambie)
    python actualizar_todo.py --solo-bot (solo descarga vacaciones, no procesa ni corre maestro)
    python actualizar_todo.py --solo-si-no-corrio-hoy   (sale de inmediato si ya se completo hoy)
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
BOT_MAESTRO = os.path.join(AQUI, "bot_maestro.py")
ESTADO = os.path.join(AQUI, "estado_bot.json")
LOG_DIR = os.path.join(AQUI, "logs")
os.makedirs(LOG_DIR, exist_ok=True)

PYEXE = sys.executable
CONFIG_BOT = os.path.join(AQUI, "config_bot.json")
PATRON = "VACRptMotivo_*.xlsx"
_POPEN_FLAGS = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0


def log(msg: str):
    linea = f"[{datetime.datetime.now():%Y-%m-%d %H:%M:%S}] {msg}"
    print(linea, flush=True)
    with open(os.path.join(LOG_DIR, f"actualizar_{datetime.date.today():%Y%m%d}.log"),
              "a", encoding="utf-8") as f:
        f.write(linea + "\n")


def _carpeta_descarga() -> str:
    """Misma resolucion que bot_adryan.py: la carpeta_descarga de config_bot.json
    si existe en ESTA maquina, si no el Downloads del usuario actual. No se puede
    asumir una ruta fija (config_bot.json trae la de la PC donde se configuro
    originalmente, que no es necesariamente esta)."""
    c = ""
    try:
        with open(CONFIG_BOT, encoding="utf-8") as f:
            c = (json.load(f).get("carpeta_descarga", "") or "").strip()
    except Exception:
        pass
    if c and os.path.isdir(c):
        return c
    return os.path.join(os.path.expanduser("~"), "Downloads")


def crudo_mas_reciente():
    archivos = glob.glob(os.path.join(_carpeta_descarga(), PATRON))
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
    solo_si_no_corrio_hoy = "--solo-si-no-corrio-hoy" in sys.argv

    estado = leer_estado()
    hoy = datetime.date.today().isoformat()
    if solo_si_no_corrio_hoy and estado.get("ultima_corrida_diaria_completa", "")[:10] == hoy:
        log(f"Ya se completo la actualizacion diaria hoy ({hoy}); no se repite.")
        return 0

    # 1) Descargar vacaciones
    log("=== Paso 1: descargar crudo de Adryan (vacaciones) ===")
    rc = subprocess.run([PYEXE, os.path.join(AQUI, "bot_adryan.py")], creationflags=_POPEN_FLAGS).returncode
    if rc != 0:
        log(f"El bot de vacaciones fallo (rc={rc}). Se intentara igual el maestro mas abajo.")
    else:
        crudo = crudo_mas_reciente()
        if not crudo:
            log("No se encontro ningun crudo descargado.")
        else:
            log(f"Crudo mas reciente: {os.path.basename(crudo)}")
            if solo_bot:
                log("Modo --solo-bot: no se corre el pipeline.")
            else:
                # 2) Detectar cambios
                nuevo_hash = hash_contenido(crudo)
                if not forzar and estado.get("ultimo_hash") == nuevo_hash:
                    log("Los datos de vacaciones NO cambiaron desde la ultima corrida -> NO se corre el pipeline.")
                    estado["ultima_revision"] = datetime.datetime.now().isoformat(timespec="seconds")
                else:
                    # 3) Correr pipeline
                    log("=== Paso 2: datos nuevos -> corriendo pipeline ===")
                    rc_pipe = subprocess.run([PYEXE, MOTOR], cwd=os.path.dirname(MOTOR), creationflags=_POPEN_FLAGS).returncode
                    if rc_pipe != 0:
                        log(f"El pipeline termino con codigo {rc_pipe}.")
                    else:
                        estado["ultimo_hash"] = nuevo_hash
                        estado["ultimo_crudo"] = os.path.basename(crudo)
                        estado["ultima_actualizacion"] = datetime.datetime.now().isoformat(timespec="seconds")

    # 4) Descargar y sanitizar el Maestro de Personal (siempre, salvo --solo-bot)
    rc_maestro = None
    if not solo_bot:
        if os.path.isfile(BOT_MAESTRO):
            log("=== Paso 3: descargando Maestro de Personal ===")
            rc_maestro = subprocess.run([PYEXE, BOT_MAESTRO], cwd=AQUI, creationflags=_POPEN_FLAGS).returncode
            if rc_maestro != 0:
                log(f"El bot del maestro termino con codigo {rc_maestro}.")
            else:
                estado["ultima_actualizacion_maestro"] = datetime.datetime.now().isoformat(timespec="seconds")
        else:
            log("bot_maestro.py no encontrado; se omite la actualizacion del maestro.")

    # Se marca "corrida diaria completa" si vacaciones no fallo (rc del paso 1
    # ya evaluado arriba via estado) y el maestro no fallo (o no se intento).
    if rc_maestro in (0, None):
        estado["ultima_corrida_diaria_completa"] = datetime.datetime.now().isoformat(timespec="seconds")

    guardar_estado(estado)
    log("=== LISTO: actualizacion diaria (vacaciones + maestro) completada. ===")
    return 0 if rc_maestro in (0, None) else rc_maestro


if __name__ == "__main__":
    sys.exit(main())
