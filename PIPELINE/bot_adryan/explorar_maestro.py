# -*- coding: utf-8 -*-
"""
Abre Adryan en modo VISIBLE para explorar la ruta del Maestro de Personal.
Se loguea automaticamente y luego espera a que navegues.
Cuando termines, cierra la ventana o presiona Enter en la consola.

Uso:
    python explorar_maestro.py
"""
import os
import sys
import json
import time
import datetime

from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

AQUI = os.path.dirname(os.path.abspath(__file__))
CONFIG = os.path.join(AQUI, "config_bot.json")

import guardar_password


def log(msg: str):
    print(f"[{datetime.datetime.now():%H:%M:%S}] {msg}", flush=True)


def main():
    with open(CONFIG, encoding="utf-8") as f:
        cfg = json.load(f)

    try:
        password = guardar_password.cargar()
    except Exception as e:
        log(f"ERROR: no hay contrasena guardada. Corre guardar_password.py primero. ({e})")
        return 1

    with sync_playwright() as pw:
        browser = pw.chromium.launch(channel=cfg["canal_navegador"], headless=False)
        sesion = os.path.join(AQUI, cfg["archivo_sesion"])
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

            # Verificar si esta logueado
            try:
                necesita_login = page.get_by_role("textbox", name="Usuario Usuario").is_visible(timeout=4000)
            except PWTimeout:
                necesita_login = False
            except Exception:
                necesita_login = False

            if necesita_login:
                log("Sesion no valida -> iniciando login...")
                page.goto(cfg["url_login"], wait_until="domcontentloaded")
                page.get_by_role("textbox", name="Usuario Usuario").fill(cfg["usuario"])
                page.get_by_role("textbox", name="Contraseña Nueva Contraseña").fill(password)
                page.get_by_role("button", name="INICIAR SESIÓN").click()
                page.wait_for_load_state("networkidle")
                context.storage_state(path=sesion)
                log("Login completado y sesion guardada.")
            else:
                log("Sesion valida reutilizada.")

            log("")
            log("=" * 60)
            log("  ADRYAN ABIERTO - NAVEGA AL MAESTRO DE PERSONAL")
            log("  Muestra los pasos para descargar el maestro.")
            log("  Cuando termines, presiona ENTER aqui para cerrar.")
            log("=" * 60)
            log("")

            input(">>> Presiona ENTER para cerrar el navegador... ")

            # Guardar sesion actualizada
            context.storage_state(path=sesion)
            log("Sesion guardada.")

        except Exception as e:
            log(f"ERROR: {e}")
            import traceback
            traceback.print_exc()
            try:
                input(">>> Presiona ENTER para cerrar... ")
            except Exception:
                pass
            return 1
        finally:
            context.close()
            browser.close()

    return 0


if __name__ == "__main__":
    sys.exit(main())
