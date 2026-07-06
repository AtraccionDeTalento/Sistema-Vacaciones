# -*- coding: utf-8 -*-
"""
enviar_cola_outlook.py
Envía todos los correos pendientes usando Outlook de escritorio.
No necesita SMTP ni contraseñas — usa la sesión de Outlook que ya tienes abierta.

Procesa in/ (pendientes) y errores/ (reintentos de PA fallidos).

Uso:
    python enviar_cola_outlook.py            # envia in/ y errores/
    python enviar_cola_outlook.py --dry-run  # muestra qué enviaría sin enviar nada
    python enviar_cola_outlook.py --solo-in  # solo procesa in/ (sin tocar errores/)
    python enviar_cola_outlook.py --revisar  # abre borradores en Outlook (no envía;
                                             # el usuario revisa y presiona Enviar)
"""
import os
import sys
import json
import shutil

# Forzar UTF-8 en la consola para evitar errores con emojis/unicode
if sys.stdout and hasattr(sys.stdout, 'reconfigure'):
    try: sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    except Exception: pass

AQUI        = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(AQUI, 'pa_config.json')
DRY_RUN     = '--dry-run' in sys.argv
SOLO_IN     = '--solo-in' in sys.argv
REVISAR     = '--revisar' in sys.argv   # abre borradores para revision manual en vez de enviar


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
    raise FileNotFoundError('No se encontró la carpeta alertas_cola. Verifica alertas_cola_dir en pa_config.json')


def enviar_via_outlook(para, nombre, asunto, html_body):
    """Envía (o abre como borrador con --revisar) usando la instancia de Outlook
    ya abierta (COM). No necesita SMTP."""
    import win32com.client as win32
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)   # 0 = MailItem
    mail.To      = para
    mail.Subject = asunto
    mail.HTMLBody = html_body
    if DRY_RUN:
        return True
    if REVISAR:
        mail.Display()   # el usuario revisa y presiona Enviar desde su cuenta
    else:
        mail.Send()
    return True


def procesar_directorio(src_dir, proc_dir, err_dir, etiqueta=''):
    """Procesa todos los .json en src_dir. Devuelve (total_ok, total_err)."""
    if not os.path.isdir(src_dir):
        return 0, 0

    archivos = sorted(f for f in os.listdir(src_dir) if f.lower().endswith('.json'))
    if not archivos:
        print(f'[OK] No hay archivos en {etiqueta or src_dir}')
        return 0, 0

    total_ok, total_err = 0, 0
    prefijo_run = '(DRY-RUN) ' if DRY_RUN else ''
    print(f'[INFO] {prefijo_run}Procesando {len(archivos)} archivo(s) de {etiqueta}...')

    for nombre_archivo in archivos:
        src = os.path.join(src_dir, nombre_archivo)
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

                p = '[DRY] ' if DRY_RUN else ''
                print(f'  {p}Enviando a {para} ({nombre_dest[:40]})')
                print(f'  Asunto: {asunto[:70]}')

                try:
                    enviar_via_outlook(para, nombre_dest, asunto, html)
                    ok_count += 1
                    total_ok += 1
                    if DRY_RUN:
                        print('  [OK] Simularia envio')
                    elif REVISAR:
                        print('  [OK] Borrador abierto en Outlook (revisa y presiona Enviar)')
                    else:
                        print('  [OK] Enviado OK')
                except Exception as e:
                    err_count += 1
                    total_err += 1
                    print(f'  [ERR] {e}')

            destino = proc_dir if err_count == 0 else err_dir
            if not DRY_RUN:
                shutil.move(src, os.path.join(destino, nombre_archivo))
                print(f'  -> Movido a {os.path.basename(destino)}/')
        except Exception as e:
            total_err += 1
            print(f'  [ERR] Error leyendo archivo: {e}')

    return total_ok, total_err


def main():
    cfg       = cargar_config()
    cola_dir  = resolver_cola_dir(cfg)
    in_dir    = os.path.join(cola_dir, 'in')
    proc_dir  = os.path.join(cola_dir, 'procesados')
    err_dir   = os.path.join(cola_dir, 'errores')

    os.makedirs(proc_dir, exist_ok=True)
    os.makedirs(err_dir,  exist_ok=True)

    total_ok, total_err = 0, 0

    # Primero procesa in/ (pendientes normales)
    ok, err = procesar_directorio(in_dir, proc_dir, err_dir, 'in/ (pendientes)')
    total_ok += ok
    total_err += err

    # Luego procesa errores/ (reintentos de Power Automate fallidos), salvo --solo-in
    if not SOLO_IN:
        ok, err = procesar_directorio(err_dir, proc_dir, err_dir, 'errores/ (reintentos)')
        total_ok += ok
        total_err += err

    print(f'\n{"="*50}')
    etiqueta = 'borradores abiertos' if REVISAR and not DRY_RUN else 'enviados'
    print(f'RESUMEN {"(DRY-RUN) " if DRY_RUN else ""}: {total_ok} {etiqueta}, {total_err} errores')
    if DRY_RUN:
        print('(Nada fue enviado -- usa sin --dry-run para enviar de verdad)')
    elif REVISAR and total_ok:
        print('(Los archivos se movieron a procesados/ -- los borradores quedaron abiertos en Outlook;')
        print(' revisa cada ventana y presiona Enviar. Si cierras un borrador sin enviar, se descarta.)')


if __name__ == '__main__':
    main()
