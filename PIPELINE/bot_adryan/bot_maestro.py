# -*- coding: utf-8 -*-
"""
Bot que descarga el Maestro de Personal desde Adryan y lo sanitiza.

Pasos:
  1. Login en Adryan (reutiliza sesion guardada)
  2. Personal > Maestro del Personal sin filtro > descargar
  3. Sanitizar el Excel descargado:
     - Eliminar los 3 primeros registros (dueños de la corporación)
     - Sustituir TODOS los emails (col DIRECCIÓN ELECTRÓNICA) por jlopezp@usil.edu.pe

Uso:
    python bot_maestro.py            (headless según config)
    python bot_maestro.py --visible  (fuerza ver el navegador)
"""
import os
import sys
import json
import time
import shutil
import datetime
import traceback

from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

AQUI = os.path.dirname(os.path.abspath(__file__))
CONFIG = os.path.join(AQUI, "config_bot.json")
LOG_DIR = os.path.join(AQUI, "logs")
os.makedirs(LOG_DIR, exist_ok=True)

import guardar_password

FILAS_DUENOS = 3
EMAIL_SEGURO = "jlopezp@usil.edu.pe"
COL_EMAIL = 20  # columna "DIRECCIÓN ELECTRÓNICA" (0-based en la data, fila 11 = headers)
FILA_HEADER = 11
FILA_DATA_INICIO = 12


def log(msg: str):
    linea = f"[{datetime.datetime.now():%Y-%m-%d %H:%M:%S}] {msg}"
    print(linea, flush=True)
    with open(os.path.join(LOG_DIR, f"maestro_{datetime.date.today():%Y%m%d}.log"),
              "a", encoding="utf-8") as f:
        f.write(linea + "\n")


def cargar_config() -> dict:
    with open(CONFIG, encoding="utf-8") as f:
        return json.load(f)


def esta_logueado(page) -> bool:
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


def _navegar_menu_maestro(page):
    """Fallback: navegar por el menu lateral icono Personal > Maestro del Personal."""
    page.goto("https://adryancloudusil.sapia.com.pe/Home/IndexAdminDashboard",
              wait_until="domcontentloaded")
    page.wait_for_load_state("networkidle")
    time.sleep(2)

    # Los iconos del menu lateral no tienen aria-label; buscamos por titulo/href
    found = page.evaluate("""() => {
        const links = document.querySelectorAll('a[title], a[data-original-title]');
        for (const a of links) {
            const t = (a.title || a.getAttribute('data-original-title') || '').toLowerCase();
            if (t.includes('personal')) { a.click(); return 'clicked-title'; }
        }
        // Fallback: buscar por texto del submenu
        const all = document.querySelectorAll('.sidebar a, .menu-lateral a, nav a, #sidebar a');
        for (const a of all) {
            const t = (a.textContent || '').trim().toLowerCase();
            if (t === 'personal') { a.click(); return 'clicked-text'; }
        }
        // Ultimo intento: segundo icono del menu lateral
        const icons = document.querySelectorAll('.sidebar-menu > li > a, .nav-sidebar > li > a');
        if (icons.length >= 2) { icons[1].click(); return 'clicked-index'; }
        return 'not-found';
    }""")
    log(f"Menu Personal: {found}")
    time.sleep(2)

    # Buscar submenu "Maestro del Personal"
    page.get_by_text("Maestro del Personal", exact=False).first.click()
    page.wait_for_load_state("networkidle")
    time.sleep(3)
    log("Maestro abierto por menu lateral.")


def descargar_maestro(page, cfg) -> str:
    log("Navegando a Maestro del Personal...")

    # Intentar URL directa primero (más robusto que navegar por menú)
    url_maestro = cfg.get("url_maestro", "").strip()
    if url_maestro:
        page.goto(url_maestro, wait_until="domcontentloaded")
        page.wait_for_load_state("networkidle")
        time.sleep(2)
        if page.locator(".ctn-iconos-edit-tabla > a").first.is_visible(timeout=5000):
            log("Maestro abierto por URL directa.")
        else:
            log("URL directa no mostro el maestro; intento por menu lateral.")
            _navegar_menu_maestro(page)
    else:
        _navegar_menu_maestro(page)

    log("Descargando maestro...")

    # La descarga se dispara via JS (reportpersonallistjs.export(0)),
    # no por navegacion. Usamos expect_download + evaluate para capturarla.
    try:
        with page.expect_download(timeout=cfg["timeout_ms"]) as dl_info:
            page.evaluate("reportpersonallistjs.export(0)")
        download = dl_info.value
    except PWTimeout:
        # Fallback: intentar con click + no_wait_after
        log("evaluate no disparo descarga, reintentando con click...")
        with page.expect_download(timeout=cfg["timeout_ms"]) as dl_info:
            page.locator(".ctn-iconos-edit-tabla > a").first.click(no_wait_after=True)
        download = dl_info.value

    nombre = download.suggested_filename or f"PersonalMaestro_{datetime.datetime.now():%m_%d_%Y %H_%M_%S}.xlsx"
    destino = os.path.join(cfg["carpeta_descarga"], nombre)
    download.save_as(destino)
    log(f"Descargado: {destino}")
    return destino


def sanitizar_maestro(ruta_xlsx: str) -> str:
    """Elimina los 3 primeros registros (dueños) y reemplaza todos los emails."""
    import openpyxl

    log(f"Sanitizando {os.path.basename(ruta_xlsx)}...")
    wb = openpyxl.load_workbook(ruta_xlsx)
    ws = wb[wb.sheetnames[0]]

    # 1) Eliminar los 3 primeros registros de data (filas 12, 13, 14)
    for i in range(FILAS_DUENOS):
        ws.delete_rows(FILA_DATA_INICIO)
    log(f"  Eliminados {FILAS_DUENOS} registros de dueños.")

    # 2) Sustituir todos los emails por el email seguro
    col_idx = COL_EMAIL + 1  # openpyxl usa 1-based
    emails_reemplazados = 0
    for row in ws.iter_rows(min_row=FILA_DATA_INICIO, min_col=col_idx, max_col=col_idx):
        cell = row[0]
        if cell.value and "@" in str(cell.value):
            cell.value = EMAIL_SEGURO
            emails_reemplazados += 1
    log(f"  Emails reemplazados: {emails_reemplazados} -> {EMAIL_SEGURO}")

    # Guardar
    sanitizado = ruta_xlsx.replace(".xlsx", "_sanitizado.xlsx")
    wb.save(sanitizado)
    wb.close()
    log(f"  Archivo sanitizado: {sanitizado}")

    wb2 = openpyxl.load_workbook(sanitizado, read_only=True, data_only=True)
    ws2 = wb2[wb2.sheetnames[0]]
    total = sum(1 for row in ws2.iter_rows(min_row=FILA_DATA_INICIO, values_only=True) if row and row[0])
    wb2.close()
    log(f"  Registros finales: {total}")

    return sanitizado


def main():
    visible = "--visible" in sys.argv
    cfg = cargar_config()
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
            log("Verificando sesion...")
            page.goto("https://adryancloudusil.sapia.com.pe/Home/IndexAdminDashboard",
                       wait_until="domcontentloaded")
            if not esta_logueado(page):
                hacer_login(page, cfg, password)
                context.storage_state(path=sesion)
                log("Sesion guardada/actualizada.")
            else:
                log("Sesion valida reutilizada.")

            destino = descargar_maestro(page, cfg)
            context.storage_state(path=sesion)

            sanitizado = sanitizar_maestro(destino)
            log("OK - maestro descargado y sanitizado.")
            print(f"ARCHIVO_MAESTRO={sanitizado}")
            return 0

        except Exception as e:
            log(f"ERROR: {e}")
            log(traceback.format_exc())
            try:
                shot = os.path.join(LOG_DIR, f"error_maestro_{datetime.datetime.now():%Y%m%d_%H%M%S}.png")
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
