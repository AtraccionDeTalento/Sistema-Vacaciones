# -*- coding: utf-8 -*-
"""
Diagnostico de UN SOLO intento de login a Adryan (sin reintentos).

Proposito: cuando el bot normal (bot_adryan.py / bot_maestro.py) falla el login
de forma rara (usuario se queda escrito, contrasena vuelve a quedar vacia, sin
toast de error), correr muchos reintentos automaticos empeora las cosas si Adryan
esta limitando/bloqueando por volumen de intentos. Este script hace UN solo
intento, deja el navegador visible y abierto para que lo mires, guarda captura +
HTML de la pagina, y muestra tu IP publica (para comparar si el bloqueo es por
red/IP entre distintas computadoras).

Uso:
    python diagnostico_login.py
"""
import os
import sys
import json
import datetime
import urllib.request

from playwright.sync_api import sync_playwright

import bot_adryan  # reusa _campo_visible, cargar_config, en_login, log

AQUI = os.path.dirname(os.path.abspath(__file__))
LOG_DIR = os.path.join(AQUI, "logs")
os.makedirs(LOG_DIR, exist_ok=True)


def _ip_publica():
    try:
        with urllib.request.urlopen("https://api.ipify.org", timeout=5) as r:
            return r.read().decode().strip()
    except Exception as e:
        return f"(no se pudo obtener: {e})"


def main():
    print("=" * 70)
    print(" DIAGNOSTICO DE LOGIN A ADRYAN - UN SOLO INTENTO")
    print("=" * 70)
    print(f" Hora: {datetime.datetime.now():%Y-%m-%d %H:%M:%S}")
    print(f" IP publica de esta maquina/red: {_ip_publica()}")

    cfg = bot_adryan.cargar_config()
    try:
        import guardar_password
        password = guardar_password.cargar()
    except Exception as e:
        print(f" ERROR: no se pudo cargar la contrasena guardada: {e}")
        return 2
    print(f" Usuario: {cfg['usuario']}   Contrasena cargada: {'SI (len=' + str(len(password)) + ')' if password else 'NO'}")
    print("-" * 70)
    print(" Se abrira Chrome visible. NO se hara ningun reintento automatico.")
    print(" Observa la pantalla: si aparece un mensaje de error, anotalo.")
    print("-" * 70)

    with sync_playwright() as pw:
        browser = pw.chromium.launch(channel=cfg["canal_navegador"], headless=False)
        context = browser.new_context()
        context.set_default_timeout(cfg["timeout_ms"])
        page = context.new_page()

        page.goto(cfg["url_login"], wait_until="domcontentloaded")
        page.wait_for_load_state("networkidle")

        campo_usuario = bot_adryan._campo_visible(page, "Ingrese Usuario", None, "Usuario Usuario")
        campo_clave = bot_adryan._campo_visible(page, "Ingrese Contraseña", None, "Contraseña Nueva Contraseña")

        campo_usuario.click()
        campo_usuario.fill("")
        campo_usuario.fill(cfg["usuario"])
        page.wait_for_timeout(300)

        campo_clave.click()
        campo_clave.fill("")
        campo_clave.fill(password)
        page.wait_for_timeout(300)

        valor_usuario = campo_usuario.input_value()
        valor_clave = campo_clave.input_value()
        print(f" Campo usuario justo antes de enviar: '{valor_usuario}' (esperado '{cfg['usuario']}')")
        print(f" Campo contrasena justo antes de enviar: {'OK, coincide' if valor_clave == password else 'NO COINCIDE (esto ya explicaria el fallo)'}")

        page.get_by_role("button", name="INICIAR SESIÓN").click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(2000)

        sello = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        shot = os.path.join(LOG_DIR, f"diagnostico_{sello}.png")
        html = os.path.join(LOG_DIR, f"diagnostico_{sello}.html")
        page.screenshot(path=shot, full_page=True)
        with open(html, "w", encoding="utf-8") as f:
            f.write(page.content())

        sigue_en_login = bot_adryan.en_login(page)
        print("-" * 70)
        if sigue_en_login:
            print(" RESULTADO: FALLO - el formulario de login sigue en pantalla.")
        else:
            print(" RESULTADO: OK - el login funciono, ya no esta en la pantalla de login.")
        print(f" Captura: {shot}")
        print(f" HTML de la pagina: {html}")
        print("=" * 70)

        input(" Presiona ENTER para cerrar el navegador y terminar...")
        context.close()
        browser.close()

    return 0


if __name__ == "__main__":
    sys.exit(main())
