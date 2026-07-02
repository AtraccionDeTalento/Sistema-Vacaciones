# -*- coding: utf-8 -*-
"""
Guarda la contrasena de Adryan.
Primario: DPAPI de Windows (cifrado por maquina).
Fallback: pa_config.json -> adryan_password (util en distribuciones sin DPAPI previo).
"""
import os
import sys
import json
import getpass

AQUI = os.path.dirname(os.path.abspath(__file__))
ARCHIVO_CRED = os.path.join(AQUI, "cred_adryan.bin")
# pa_config.json esta 2 niveles arriba (raiz del proyecto)
PA_CONFIG = os.path.join(AQUI, '..', '..', 'pa_config.json')


def _pa_config_leer():
    try:
        with open(PA_CONFIG, encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return {}


def _pa_config_escribir(patch: dict):
    try:
        cfg = _pa_config_leer()
        cfg.update(patch)
        with open(PA_CONFIG, 'w', encoding='utf-8') as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def guardar(password: str) -> None:
    """Guarda en DPAPI Y en pa_config.json para que funcione en cualquier maquina."""
    # Guardar en pa_config.json siempre (accesible sin DPAPI)
    _pa_config_escribir({'adryan_password': password})

    # Intentar guardar en DPAPI tambien (mas seguro si esta disponible)
    try:
        import win32crypt
        blob = win32crypt.CryptProtectData(
            password.encode("utf-8"), "adryan_bot", None, None, None, 0
        )
        with open(ARCHIVO_CRED, "wb") as f:
            f.write(blob)
    except Exception:
        pass  # Si no hay pywin32, pa_config.json es suficiente


def cargar() -> str:
    """Intenta DPAPI primero; si falla, lee pa_config.json."""
    # 1) DPAPI
    try:
        import win32crypt
        with open(ARCHIVO_CRED, "rb") as f:
            blob = f.read()
        _, datos = win32crypt.CryptUnprotectData(blob, None, None, None, 0)
        return datos.decode("utf-8")
    except Exception:
        pass

    # 2) pa_config.json
    pw = _pa_config_leer().get('adryan_password', '').strip()
    if pw:
        return pw

    raise RuntimeError("no hay contrasena guardada. Configurala desde el dashboard.")


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "--verificar":
        try:
            pw = cargar()
            print(f"OK: contrasena guardada (longitud {len(pw)}).")
        except Exception as e:
            print(f"ERROR: {e}")
        sys.exit(0)

    print("Pega o escribe la contrasena de Adryan (no se mostrara en pantalla):")
    pw = getpass.getpass("Contrasena: ")
    if not pw:
        print("No se ingreso nada. Cancelado.")
        sys.exit(1)
    guardar(pw)
    print(f"Listo. Contrasena guardada.")
    print("Verificando...")
    print("OK" if cargar() == pw else "ERROR: no coincide")
