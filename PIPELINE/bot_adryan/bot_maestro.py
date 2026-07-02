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
    """Navegar por el menu: Personal > Maestro del Personal sin filtro."""
    page.goto("https://adryancloudusil.sapia.com.pe/Home/IndexAdminDashboard",
              wait_until="domcontentloaded")
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(2000)

    # Clic en el item de menu "Personal" (igual que el script grabado)
    try:
        page.get_by_role("link", name="Personal").first.click()
        page.wait_for_timeout(1500)
    except Exception:
        # Fallback JS
        page.evaluate("""() => {
            const all = [...document.querySelectorAll('a')];
            const a = all.find(el => (el.textContent || '').trim() === 'Personal');
            if (a) a.click();
        }""")
        page.wait_for_timeout(1500)

    # Clic en "Maestro del Personal sin [filtro]"
    try:
        page.get_by_role("link", name="Maestro del Personal sin").click()
    except Exception:
        # Fallback: buscar por texto parcial
        page.get_by_role("link", name="Maestro del Personal").first.click()

    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(3000)
    log("Maestro abierto.")


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
    timeout_dl = max(cfg.get("timeout_ms", 60000), 120000)  # minimo 2 minutos

    # La descarga se dispara via JS (reportpersonallistjs.export(0)),
    # no por navegacion. Usamos expect_download + evaluate para capturarla.
    try:
        with page.expect_download(timeout=timeout_dl) as dl_info:
            page.evaluate("reportpersonallistjs.export(0)")
        download = dl_info.value
    except PWTimeout:
        # Fallback: intentar con click en el boton de descarga
        log("evaluate no disparo descarga, reintentando con click...")
        with page.expect_download(timeout=timeout_dl) as dl_info:
            page.locator(".ctn-iconos-edit-tabla > a").first.click(no_wait_after=True)
        download = dl_info.value

    nombre = download.suggested_filename or f"PersonalMaestro_{datetime.datetime.now():%m_%d_%Y %H_%M_%S}.xlsx"
    carpeta = cfg.get("carpeta_descarga", "").strip()
    if not carpeta or not os.path.isdir(carpeta):
        carpeta = os.path.join(os.path.expanduser("~"), "Downloads")
        os.makedirs(carpeta, exist_ok=True)
        log(f"carpeta_descarga no encontrada; usando {carpeta}")
    destino = os.path.join(carpeta, nombre)
    download.save_as(destino)
    log(f"Descargado: {destino}")
    return destino


def sanitizar_maestro(ruta_xlsx: str) -> str:
    """Elimina los 3 primeros registros (dueños) y reemplaza todos los emails."""
    import openpyxl

    log(f"Sanitizando {os.path.basename(ruta_xlsx)}...")
    wb = openpyxl.load_workbook(ruta_xlsx)
    ws = wb[wb.sheetnames[0]]

    # Leer todas las filas en memoria para procesarlas rápidamente
    all_rows = []
    for row in ws.iter_rows(values_only=True):
        all_rows.append(row)

    # El encabezado está en la fila 11 (índice 10)
    # Los datos comienzan en la fila 12 (índice 11)
    metadata_and_header = all_rows[:FILA_HEADER] # 11 filas (0 a 10)
    data_rows = all_rows[FILA_HEADER:]          # Desde fila 12 (índice 11)

    # 1) Eliminar los 3 primeros registros de data
    filtered_data_rows = data_rows[FILAS_DUENOS:]
    log(f"  Eliminados {FILAS_DUENOS} registros de dueños (en memoria).")

    # 2) Sustituir todos los emails por el email seguro
    emails_reemplazados = 0
    sanitized_data = []
    for row in filtered_data_rows:
        if not row:
            continue
        row_list = list(row)
        if len(row_list) > COL_EMAIL:
            val = row_list[COL_EMAIL]
            if val and "@" in str(val):
                row_list[COL_EMAIL] = EMAIL_SEGURO
                emails_reemplazados += 1
        sanitized_data.append(row_list)
    log(f"  Emails reemplazados: {emails_reemplazados} -> {EMAIL_SEGURO}")

    # Recrear la hoja para evitar la lentitud extrema de delete_rows
    sheet_title = ws.title
    new_ws = wb.create_sheet(title="SanitizadoTemp")

    # Escribir cabecera y metadata
    for r_idx, row_val in enumerate(metadata_and_header, start=1):
        for c_idx, val in enumerate(row_val, start=1):
            new_ws.cell(row=r_idx, column=c_idx, value=val)

    # Escribir filas de datos sanitizadas
    for r_idx, row_val in enumerate(sanitized_data, start=FILA_DATA_INICIO):
        for c_idx, val in enumerate(row_val, start=1):
            new_ws.cell(row=r_idx, column=c_idx, value=val)

    # Eliminar hoja antigua y renombrar la nueva
    wb.remove(ws)
    new_ws.title = sheet_title

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
