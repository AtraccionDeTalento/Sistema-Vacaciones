# -*- coding: utf-8 -*-
"""
Guarda la contrasena de Adryan CIFRADA con DPAPI de Windows.
Solo tu usuario de Windows (en esta maquina) puede descifrarla.
Uso:
    python guardar_password.py
Te la pide por teclado (no se ve) y la escribe cifrada en cred_adryan.bin.
"""
import os
import sys
import getpass

import win32crypt  # viene con pywin32 (ya instalado para el pipeline)

AQUI = os.path.dirname(os.path.abspath(__file__))
ARCHIVO_CRED = os.path.join(AQUI, "cred_adryan.bin")


def guardar(password: str) -> None:
    blob = win32crypt.CryptProtectData(
        password.encode("utf-8"),
        "adryan_bot",      # descripcion
        None, None, None, 0
    )
    with open(ARCHIVO_CRED, "wb") as f:
        f.write(blob)


def cargar() -> str:
    with open(ARCHIVO_CRED, "rb") as f:
        blob = f.read()
    _, datos = win32crypt.CryptUnprotectData(blob, None, None, None, 0)
    return datos.decode("utf-8")


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "--verificar":
        try:
            pw = cargar()
            print(f"OK: contrasena guardada (longitud {len(pw)}).")
        except Exception as e:
            print(f"ERROR leyendo cred_adryan.bin: {e}")
        sys.exit(0)

    print("Pega o escribe la contrasena de Adryan (no se mostrara en pantalla):")
    pw = getpass.getpass("Contrasena: ")
    if not pw:
        print("No se ingreso nada. Cancelado.")
        sys.exit(1)
    guardar(pw)
    print(f"Listo. Contrasena cifrada en: {ARCHIVO_CRED}")
    print("Verificando que se pueda descifrar...")
    print("OK" if cargar() == pw else "ERROR: no coincide")
