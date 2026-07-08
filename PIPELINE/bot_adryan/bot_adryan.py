# -*- coding: utf-8 -*-
"""
Bot que descarga el crudo de vacaciones (VACRptMotivo_*.xlsx) desde Adryan.

- Login fresco de un solo intento en cada corrida (igual que diagnostico_login.py,
  confirmado funcional): no reutiliza sesion guardada ni reintenta el login.
- Se loguea con usuario + contrasena cifrada (DPAPI) y graba la sesion resultante
  solo como registro (no se vuelve a cargar en la siguiente corrida).
- Replica los pasos que grabaste: Personal -> Vacaciones por Motivo -> filtros -> Buscar -> descargar.
- Guarda el archivo en la carpeta de Descargas que lee el pipeline.

Uso:
    python bot_adryan.py            (headless segun config)
    python bot_adryan.py --visible  (fuerza ver el navegador)
"""
import os
import sys
import json
import datetime
import traceback

try:
    from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout
except ImportError:
    print("ERROR: playwright no esta instalado en este entorno de Python.")
    print("Esta funcion solo esta disponible desde la maquina administrativa.")
    print("Para instalar: pip install playwright && playwright install chromium")
    sys.exit(1)

AQUI = os.path.dirname(os.path.abspath(__file__))
CONFIG = os.path.join(AQUI, "config_bot.json")
LOG_DIR = os.path.join(AQUI, "logs")
os.makedirs(LOG_DIR, exist_ok=True)

import guardar_password  # para leer la contrasena cifrada


class SesionCaducada(RuntimeError):
    """Adryan redirigio al login a mitad del flujo: hay que re-loguear y reintentar."""


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


def en_login(page) -> bool:
    """True si la pagina actual es (o redirigio a) la pantalla de login de Adryan."""
    try:
        return page.get_by_role("textbox", name="Usuario Usuario").is_visible(timeout=1500)
    except Exception:
        return False


def _campo_visible(page, locator_placeholder, locator_role, nombre_role):
    """Adryan tiene un campo oculto (modal de cambio de clave) con el mismo
    accessible-name que el campo visible de login, asi que get_by_role a veces
    apunta al elemento equivocado. Se prioriza el placeholder, que es unico
    al campo que realmente se ve en pantalla, y solo se cae al role si no
    hay match por placeholder."""
    por_placeholder = page.get_by_placeholder(locator_placeholder)
    try:
        if por_placeholder.count() > 0 and por_placeholder.first.is_visible():
            return por_placeholder.first
    except Exception:
        pass
    return page.get_by_role("textbox", name=nombre_role)


def _intentar_login(page, cfg, password):
    campo_usuario = _campo_visible(page, "Ingrese Usuario", None, "Usuario Usuario")
    campo_clave = _campo_visible(page, "Ingrese Contraseña", None, "Contraseña Nueva Contraseña")

    # Igual que el paso grabado: click antes de fill (asegura foco/bind de Angular
    # antes de escribir) y una pausa corta para que el framework registre el value
    # antes de enviar el formulario. Bajo carga de sistema (CPU ocupada, muchos
    # procesos), el submit a veces viaja con los campos vacios porque Angular
    # todavia no termino de bindear cuando se hace click en INICIAR SESION.
    campo_usuario.click()
    campo_usuario.fill("")
    campo_usuario.fill(cfg["usuario"])
    page.wait_for_timeout(400)

    campo_clave.click()
    campo_clave.fill("")
    campo_clave.fill(password)
    page.wait_for_timeout(400)

    # Verificar que el value realmente quedo escrito antes de enviar. Reintentar
    # varias veces (no solo una) porque bajo carga una sola pasada puede no alcanzar.
    for _ in range(3):
        if campo_usuario.input_value() == cfg["usuario"] and campo_clave.input_value() == password:
            break
        log("Los campos no registraron el valor escrito; reintentando fill...")
        campo_usuario.fill(cfg["usuario"])
        campo_clave.fill(password)
        page.wait_for_timeout(500)

    page.get_by_role("button", name="INICIAR SESIÓN").click()
    page.wait_for_load_state("networkidle")


def hacer_login(page, cfg, password):
    # Un solo intento, igual que diagnostico_login.py (confirmado funcional):
    # reintentar el login varias veces seguidas es justo el patron que hace
    # que Adryan limite/bloquee (deteccion anti-bot por volumen de intentos).
    log("Iniciando login (intento unico, igual al diagnostico)...")
    page.goto(cfg["url_login"], wait_until="domcontentloaded")
    page.wait_for_load_state("networkidle")

    _intentar_login(page, cfg, password)
    page.wait_for_timeout(2000)

    if en_login(page):
        raise RuntimeError(
            "Login fallido: Adryan sigue mostrando el formulario de acceso tras "
            "el intento. La clave puede estar vencida, o Adryan esta limitando "
            "temporalmente los intentos de login automatico (confirma primero "
            "si el login manual funciona antes de tocar la contrasena guardada)."
        )
    log("Login verificado OK.")


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
    log(f"Fijando fechas {cfg['fecha_inicio']} -> {cfg['fecha_termino']} via interaccion de UI...")

    # Fecha Inicio
    page.get_by_role("textbox", name="Fecha Inicio").click(force=True)
    page.wait_for_timeout(1000)  # dar tiempo a que se abra el calendario
    page.get_by_role("combobox").nth(2).select_option(cfg.get("mes_inicio", "3"))
    page.get_by_role("gridcell", name=cfg.get("dia_inicio", "01/04/")).click(force=True)

    # Pequeña pausa para que la animacion del calendario se cierre
    log("Esperando cierre del primer calendario...")
    page.wait_for_timeout(2000)

    # Fecha Término
    page.get_by_role("textbox", name="Fecha Término").click(force=True)
    page.wait_for_timeout(1000)  # dar tiempo a que se abra el segundo calendario
    page.get_by_role("combobox").nth(2).select_option(cfg.get("mes_termino", "6"))
    page.get_by_role("gridcell", name=cfg.get("dia_termino", "30/07/")).click(force=True)

    page.wait_for_timeout(1000)
    log("Fechas fijadas.")


def _carpeta_descarga(cfg) -> str:
    """Devuelve la carpeta de descarga: la del config si existe, si no la del usuario."""
    c = cfg.get("carpeta_descarga", "").strip()
    if c and os.path.isdir(c):
        return c
    fallback = os.path.join(os.path.expanduser("~"), "Downloads")
    os.makedirs(fallback, exist_ok=True)
    log(f"carpeta_descarga '{c}' no encontrada; usando {fallback}")
    return fallback


def buscar_y_descargar(page, cfg) -> str:
    log("Clic en Buscar...")
    page.get_by_role("button", name="Buscar").click()

    log("Esperando a que el boton de Excel (.mr-3) aparezca en la pagina...")
    timeout_dl = max(cfg.get("timeout_ms", 60000), 120000)

    # El paso grabado utiliza .mr-3, esperamos a que sea visible
    try:
        # Aumentamos el timeout a 60s (60000ms) para que tenga tiempo de procesar
        page.locator(".mr-3").wait_for(state="visible", timeout=60000)
    except Exception as e:
        raise RuntimeError("No se encontro el boton de descarga Excel en la pagina despues de 60s.") from e

    log("Boton detectado, iniciando descarga...")
    try:
        with page.expect_download(timeout=timeout_dl) as dl_info:
            page.locator(".mr-3").first.click()
        download = dl_info.value
    except PWTimeout:
        # Caso tipico: Adryan invalido la sesion a mitad del flujo y la pagina
        # redirigio al login, por eso la descarga nunca llega.
        if en_login(page):
            raise SesionCaducada(
                "Adryan expulso la sesion al hacer la descarga (redirigio al login)."
            )
        raise

    carpeta = _carpeta_descarga(cfg)
    nombre = download.suggested_filename or f"{cfg['patron_crudo']}_{datetime.datetime.now():%m_%d_%Y %H_%M_%S}.xlsx"
    destino = os.path.join(carpeta, nombre)
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
        # Los args evitan que Chrome muestre ventanas fantasma en algunos entornos Windows
        browser = pw.chromium.launch(
            channel=cfg["canal_navegador"],
            headless=headless,
            args=["--disable-gpu", "--window-position=-32000,-32000", "--hide-scrollbars"]
        )
        # Sin storage_state: igual que diagnostico_login.py, siempre login fresco
        # directo a la pagina de login. Reutilizar sesion guardada (y verificarla
        # yendo primero al dashboard) es lo que se sospechaba causaba los fallos.
        context = browser.new_context(accept_downloads=True)
        context.set_default_timeout(cfg["timeout_ms"])
        page = context.new_page()

        try:
            hacer_login(page, cfg, password)
            context.storage_state(path=sesion)
            log("Sesion guardada/actualizada.")

            ir_al_reporte(page, cfg)
            fijar_filtros(page, cfg)
            destino = buscar_y_descargar(page, cfg)
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
