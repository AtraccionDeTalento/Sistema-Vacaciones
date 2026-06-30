# -*- coding: utf-8 -*-
"""
Bot que descarga el crudo de vacaciones (VACRptMotivo_*.xlsx) desde Adryan.

- Reusa sesion guardada (sesion_adryan.json) para no loguearse cada vez.
- Si la sesion caduco, se loguea con usuario + contrasena cifrada (DPAPI) y la regraba.
- Replica los pasos que grabaste: Personal -> Vacaciones por Motivo -> filtros -> Buscar -> descargar.
- Guarda el archivo en la carpeta de Descargas que lee el pipeline.

Uso:
    python bot_adryan.py            (headless segun config)
    python bot_adryan.py --visible  (fuerza ver el navegador)
"""
import os
import sys
import json
import time
import datetime
import traceback

from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

AQUI = os.path.dirname(os.path.abspath(__file__))
CONFIG = os.path.join(AQUI, "config_bot.json")
LOG_DIR = os.path.join(AQUI, "logs")
os.makedirs(LOG_DIR, exist_ok=True)

import guardar_password  # para leer la contrasena cifrada


def log(msg: str):
    linea = f"[{datetime.datetime.now():%Y-%m-%d %H:%M:%S}] {msg}"
    print(linea, flush=True)
    # OneDrive puede bloquear el archivo mientras lo sincroniza; reintentar 3 veces
    ruta_log = os.path.join(LOG_DIR, f"bot_{datetime.date.today():%Y%m%d}.log")
    for _intento in range(3):
        try:
            with open(ruta_log, "a", encoding="utf-8") as f:
                f.write(linea + "\n")
            break
        except PermissionError:
            if _intento < 2:
                import time as _t; _t.sleep(0.5)
            # Si sigue fallando al tercer intento, ignorar (el print ya registró el mensaje)


def cargar_config() -> dict:
    with open(CONFIG, encoding="utf-8") as f:
        return json.load(f)


def esta_logueado(page) -> bool:
    """True si NO se ve el campo de usuario (es decir, ya esta dentro)."""
    try:
        return not page.get_by_role("textbox", name="Usuario Usuario").is_visible(timeout=4000)
    except PWTimeout:
        return True
    except Exception:
        return True


def hacer_login(page, cfg, password):
    log("Sesion no valida -> iniciando login...")
    page.goto(cfg["url_login"], wait_until="domcontentloaded")
    page.get_by_role("textbox", name="Usuario Usuario").fill(cfg["usuario"])
    page.get_by_role("textbox", name="Contraseña Nueva Contraseña").fill(password)
    page.get_by_role("button", name="INICIAR SESIÓN").click()
    page.wait_for_load_state("networkidle")
    log("Login enviado.")


def ir_al_reporte(page, cfg):
    log("Navegando al reporte Vacaciones por Motivo...")
    url_directa = cfg.get("url_reporte", "").strip()
    if url_directa:
        page.goto(url_directa, wait_until="domcontentloaded")
        page.wait_for_load_state("networkidle")
        if page.get_by_role("button", name="Buscar").is_visible():
            log("Reporte abierto por URL directa.")
            return
        log("URL directa no mostro el reporte; intento por el menu lateral.")

    # Fallback: navegar por el menu lateral (icono Personal -> Vacaciones por Motivo)
    page.goto("https://adryancloudusil.sapia.com.pe/Home/IndexAdminDashboard",
              wait_until="domcontentloaded")
    page.get_by_role("link", description="Personal", exact=True).click()
    page.get_by_role("link", name="Vacaciones por Motivo").click()
    page.wait_for_load_state("networkidle")


def fijar_filtros(page, cfg):
    log(f"Fijando fechas {cfg['fecha_inicio']} -> {cfg['fecha_termino']} via pickadate...")

    def partes(fecha_ddmmaaaa: str):
        d, m, y = fecha_ddmmaaaa.split("/")
        return [int(y), int(m) - 1, int(d)]   # pickadate: [anio, mes(0=Ene), dia]

    payload = {
        "ini_id": cfg.get("id_fecha_inicio", "txtDateInitialMotiveReport"),
        "fin_id": cfg.get("id_fecha_termino", "txtDateEndMotiveReport"),
        "ini": partes(cfg["fecha_inicio"]),
        "fin": partes(cfg["fecha_termino"]),
        "motivo_id": cfg.get("id_motivo", "cboVacationMotiveReport"),
        "motivo": cfg.get("motivo_valor", "1"),
    }

    res = page.evaluate(
        """(d) => {
            const $ = window.jQuery;
            function setFecha(id, arr){
                const pk = $('#'+id).pickadate('picker');
                pk.set('select', arr);
                return pk.get('select', 'dd/mm/yyyy');
            }
            const out = { ini: setFecha(d.ini_id, d.ini), fin: setFecha(d.fin_id, d.fin) };
            const m = document.getElementById(d.motivo_id);
            if (m){ m.value = d.motivo; m.dispatchEvent(new Event('change', {bubbles:true})); out.motivo = m.value; }
            return out;
        }""",
        payload,
    )
    log(f"Filtros aplicados: {res}")


def buscar_y_descargar(page, cfg) -> str:
    log("Clic en Buscar...")
    page.get_by_role("button", name="Buscar").click()
    page.wait_for_load_state("networkidle")
    time.sleep(2)  # deja que se pinte la tabla

    log("Descargando reporte...")
    with page.expect_download(timeout=cfg["timeout_ms"]) as dl_info:
        page.locator(".mr-3").first.click()
    download = dl_info.value

    nombre = download.suggested_filename or f"{cfg['patron_crudo']}_{datetime.datetime.now():%m_%d_%Y %H_%M_%S}.xlsx"
    destino = os.path.join(cfg["carpeta_descarga"], nombre)
    download.save_as(destino)
    log(f"Descargado: {destino}")
    return destino


def parse_fecha_iso(iso_str):
    """Convierte yyyy-mm-dd a dd/mm/yyyy para Adryan."""
    parts = iso_str.split("-")
    if len(parts) == 3:
        return f"{parts[2]}/{parts[1]}/{parts[0]}"
    return iso_str


def main():
    visible = "--visible" in sys.argv
    cfg = cargar_config()

    fecha_inicio = None
    fecha_termino = None
    args = sys.argv[1:]
    for i, a in enumerate(args):
        if a == "--fecha-inicio" and i + 1 < len(args):
            fecha_inicio = args[i + 1]
        elif a == "--fecha-termino" and i + 1 < len(args):
            fecha_termino = args[i + 1]

    if fecha_inicio:
        cfg["fecha_inicio"] = parse_fecha_iso(fecha_inicio)
        log(f"Fecha inicio override: {cfg['fecha_inicio']}")
    if fecha_termino:
        cfg["fecha_termino"] = parse_fecha_iso(fecha_termino)
        log(f"Fecha termino override: {cfg['fecha_termino']}")

    headless = cfg.get("headless", True) and not visible
    sesion = os.path.join(AQUI, cfg["archivo_sesion"])

    try:
        password = guardar_password.cargar()
    except Exception as e:
        log(f"ERROR: no hay contrasena guardada. Corre guardar_password.py primero. ({e})")
        return 2

    with sync_playwright() as pw:
        browser = pw.chromium.launch(channel=cfg["canal_navegador"], headless=headless)
        ctx_args = {"accept_downloads": True}
        if os.path.exists(sesion):
            ctx_args["storage_state"] = sesion
            log("Cargando sesion guardada.")
        context = browser.new_context(**ctx_args)
        context.set_default_timeout(cfg["timeout_ms"])
        page = context.new_page()

        try:
            # 1) Verificar sesion: ir al dashboard; si redirige al login, autenticar.
            log("Verificando sesion...")
            page.goto("https://adryancloudusil.sapia.com.pe/Home/IndexAdminDashboard",
                      wait_until="domcontentloaded")
            if not esta_logueado(page):
                hacer_login(page, cfg, password)
                context.storage_state(path=sesion)  # regraba sesion fresca
                log("Sesion guardada/actualizada.")
            else:
                log("Sesion valida reutilizada.")

            # 2) Navegar al reporte y descargar.
            ir_al_reporte(page, cfg)
            fijar_filtros(page, cfg)
            destino = buscar_y_descargar(page, cfg)
            # refresca cookies por si rotaron
            context.storage_state(path=sesion)
            log("OK - descarga completada.")
            print(f"ARCHIVO_DESCARGADO={destino}")
            return 0

        except Exception as e:
            log(f"ERROR: {e}")
            log(traceback.format_exc())
            try:
                shot = os.path.join(LOG_DIR, f"error_{datetime.datetime.now():%Y%m%d_%H%M%S}.png")
                page.screenshot(path=shot, full_page=True)
                log(f"Captura del error: {shot}")
            except Exception:
                pass
            return 1
        finally:
            context.close()
            browser.close()


if __name__ == "__main__":
    sys.exit(main())
