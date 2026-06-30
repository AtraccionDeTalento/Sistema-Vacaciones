# -*- coding: utf-8 -*-
"""
enviar_cola_outlook.py
Envía todos los correos pendientes en alertas_cola/in/ usando Outlook de escritorio.
No necesita SMTP ni contraseñas — usa la sesión de Outlook que ya tienes abierta.

Uso:
    python enviar_cola_outlook.py            # envia todo lo que haya en in/
    python enviar_cola_outlook.py --dry-run  # muestra qué enviaría sin enviar nada
"""
import os
import sys
import json
import shutil
import datetime

# Forzar UTF-8 en la consola para evitar errores con emojis/unicode
if sys.stdout and hasattr(sys.stdout, 'reconfigure'):
    try: sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    except Exception: pass

AQUI        = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(AQUI, 'pa_config.json')
DRY_RUN     = '--dry-run' in sys.argv


def cargar_config():
    if os.path.isfile(CONFIG_PATH):
        with open(CONFIG_PATH, encoding='utf-8') as f:
            return json.load(f)
    return {}


def resolver_cola_dir(cfg):
    d = cfg.get('alertas_cola_dir', '').strip()
    if d and os.path.isdir(d):
        return d
    base = os.path.join(os.path.expanduser('~'),
                        'OneDrive - Universidad San Ignacio de Loyola', 'alertas_cola')
    if os.path.isdir(base):
        return base
    raise FileNotFoundError(f'No se encontró la carpeta alertas_cola. Verifica alertas_cola_dir en pa_config.json')


def enviar_via_outlook(para, nombre, asunto, html_body):
    """Envía usando la instancia de Outlook ya abierta (COM). No necesita SMTP."""
    import win32com.client as win32
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)   # 0 = MailItem
    mail.To      = para
    mail.Subject = asunto
    mail.HTMLBody = html_body
    if not DRY_RUN:
        mail.Send()
    return True


def main():
    cfg       = cargar_config()
    cola_dir  = resolver_cola_dir(cfg)
    in_dir    = os.path.join(cola_dir, 'in')
    proc_dir  = os.path.join(cola_dir, 'procesados')
    err_dir   = os.path.join(cola_dir, 'errores')

    if not os.path.isdir(in_dir):
        print(f'[WARN] Carpeta in/ no existe: {in_dir}')
        return

    os.makedirs(proc_dir, exist_ok=True)
    os.makedirs(err_dir,  exist_ok=True)

    archivos = [f for f in os.listdir(in_dir) if f.lower().endswith('.json')]
    if not archivos:
        print('[OK] No hay archivos pendientes en in/')
        return

    print(f'[INFO] {"(DRY-RUN) " if DRY_RUN else ""}Procesando {len(archivos)} archivos...')
    total_ok, total_err = 0, 0

    for nombre_archivo in sorted(archivos):
        src = os.path.join(in_dir, nombre_archivo)
        print(f'\n--- {nombre_archivo} ---')
        try:
            with open(src, 'r', encoding='utf-8') as f:
                entradas = json.load(f)
            if not isinstance(entradas, list):
                entradas = [entradas]

            ok_count, err_count = 0, 0
            for entrada in entradas:
                para        = (entrada.get('email_jefe') or entrada.get('email_destino_real') or '').strip()
                nombre_dest = entrada.get('nombre_jefe', '')
                asunto      = entrada.get('asunto', 'Alertas Vacaciones USIL')
                html        = entrada.get('mensaje_html', '')

                if not para or '@' not in para:
                    print('  [SKIP] Sin email destino')
                    continue
                if not html:
                    print('  [SKIP] Sin contenido HTML')
                    continue

                prefijo = '[DRY] ' if DRY_RUN else ''
                print(f'  {prefijo}Enviando a {para} ({nombre_dest[:40]})')
                print(f'  Asunto: {asunto[:70]}')

                try:
                    enviar_via_outlook(para, nombre_dest, asunto, html)
                    ok_count += 1
                    total_ok += 1
                    print(f'  [OK] {"Simularia envio" if DRY_RUN else "Enviado OK"}')
                except Exception as e:
                    err_count += 1
                    total_err += 1
                    print(f'  [ERR] {e}')

            # mover segun resultado real del envio
            destino = proc_dir if err_count == 0 else err_dir
            if not DRY_RUN:
                shutil.move(src, os.path.join(destino, nombre_archivo))
                print(f'  -> Movido a {os.path.basename(destino)}/')
        except Exception as e:
            total_err += 1
            print(f'  [ERR] Error leyendo archivo: {e}')

    print(f'\n{"="*50}')
    print(f'RESUMEN {"(DRY-RUN) " if DRY_RUN else ""}: {total_ok} enviados, {total_err} errores')
    if DRY_RUN:
        print('(Nada fue enviado -- usa sin --dry-run para enviar de verdad)')


if __name__ == '__main__':
    main()
