#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Buscar-Bitacoras.py
====================
Script único para consultar bitácoras de correo (Exchange Message Trace)
directamente sobre los CSV de esta carpeta, con 3 modos de búsqueda. Todos
comparten el mismo motor (multi-hilo, filtro rápido por nombre de archivo
cuando das fechas, reconocimiento del correo aunque venga con nombre de
display) y el mismo diseño de reporte Excel + PDF con formato INE.

MODO 1 — Enviados/recibidos hacia o desde dominios EXTERNOS
(omite los dominios de DOMINIOS_INTERNOS):
    python Buscar-Bitacoras.py --usuario correo@ine.mx --direccion enviados
    python Buscar-Bitacoras.py --usuario correo@ine.mx --direccion recibidos
    python Buscar-Bitacoras.py --usuario correo@ine.mx --direccion enviados --desde 2026-01-01 --hasta 2026-09-25

MODO 2 — TODO (enviados + recibidos, SIN excluir dominios internos):
    python Buscar-Bitacoras.py --usuario correo@ine.mx --todo
    python Buscar-Bitacoras.py --usuario correo@ine.mx --todo --desde 2026-06-01 --hasta 2026-06-15

MODO 3 — Búsqueda por PATRÓN (cualquier remitente -> destinatario parecido
a una palabra, o al revés con --campo remitente; tolera acentos y typos):
    python Buscar-Bitacoras.py --patron libelula
    python Buscar-Bitacoras.py --patron libelula --campo remitente
    python Buscar-Bitacoras.py --patron libelula --desde 2026-06-01 --hasta 2026-06-15

Sin argumentos, el script pregunta de forma interactiva qué modo usar.
--desde / --hasta son OPCIONALES en los 3 modos (formato AAAA-MM-DD,
horario CDMX); si no se indican, se revisa TODO el histórico disponible.

Requisitos (una sola vez):
    pip install openpyxl python-dateutil tzdata pywin32 pillow
"""

import argparse
import base64
import concurrent.futures
import csv
import difflib
import glob
import os
import re
import sys
import time
import unicodedata
from datetime import datetime, timedelta
from io import BytesIO
from zoneinfo import ZoneInfo

csv.field_size_limit(min(2**31 - 1, sys.maxsize))

try:
    from dateutil import parser as dtparser
    HAVE_DATEUTIL = True
except ImportError:
    HAVE_DATEUTIL = False

import openpyxl
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins
from openpyxl.worksheet.properties import PageSetupProperties

# ============================================================
# CONFIGURACION
# ============================================================
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CSV_DIR = r"\\VM-ADMON-WIN\E$\Bitacoras"
OUT_DIR = SCRIPT_DIR

TZ_CDMX = ZoneInfo("America/Mexico_City")
TZ_UTC = ZoneInfo("UTC")

EXCLUDE_FILES = {"fechas.csv", "usuarios.csv", "refechas.csv"}

# Dominios INTERNOS: en el modo enviados/recibidos "a externos" se OMITE
# cualquier correo hacia/desde uno de estos dominios. En el modo --todo NO
# se aplica (ahí se incluye todo, interno y externo).
DOMINIOS_INTERNOS = {
    "ine.mx",
    "sec.ife.org.mx",
    "ife.org.mx",
    "listas.ine.mx",
    "inemexico.mail.onmicrosoft.com",
}

STATUS_INFO = {
    "delivered": ("Entregado", "El mensaje llegó correctamente al buzón del destinatario."),
    "failed": ("Error", "El mensaje no pudo entregarse; fue rechazado o rebotado."),
    "pending": ("Pendiente", "El mensaje todavía está en proceso de entrega."),
    "resolved": ("Resuelto", "Notificación de un sistema de monitoreo que ya fue atendida."),
    "quarantined": ("En cuarentena", "El sistema de seguridad detuvo el mensaje por sospecha de spam o virus."),
    "filteredasspam": ("Filtrado como spam", "Se identificó como correo no deseado y no llegó a la bandeja principal."),
    "gettingstatus": ("Consultando estatus", "El sistema todavía está verificando qué pasó con el mensaje."),
    "expanded": ("Expandido", "El mensaje se envió a una lista o grupo de correo y se distribuyó a sus integrantes."),
    "recalled": ("Recall solicitado", "El remitente intentó recuperar (cancelar) este mensaje después de haberlo enviado."),
    "removedbymessagerecall": ("Retirado por recall", "El mensaje fue eliminado del buzón del destinatario porque el remitente lo recuperó a tiempo."),
    "poison": ("Mensaje dañado", "El sistema detectó un problema grave en el mensaje y lo detuvo por seguridad."),
    "none": ("Sin estatus", "No hay información de estatus registrada para este evento."),
}


def status_label(status_raw):
    info = STATUS_INFO.get((status_raw or "").strip().lower())
    return info[0] if info else (status_raw or "Sin estatus")


def status_desc(status_raw):
    info = STATUS_INFO.get((status_raw or "").strip().lower())
    return info[1] if info else "Estatus no documentado; consulte con el área de soporte si tiene dudas."


LILA_INST = "674092"
LILA_OSCURO = "49276F"
GRIS_CLARO = "F2F2F2"
BLANCO = "FFFFFF"
VERDE_OK = "E2EFDA"
ROJO_ERR = "FCE4E4"


# Logo del INE embebido en base64 (217x89 PNG) para que el script sea
# un solo archivo portable, sin depender de un archivo externo en el
# servidor.
LOGO_INE_PNG_BASE64 = (
    "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"
)


def _logo_valido():
    """
    Decodifica el logo embebido y valida que la imagen esté completa
    (no solo que el base64 decodifique). Una imagen corrupta o
    truncada suele decodificar sin error pero se ve como un cuadro
    negro/gris sólido en Excel, así que se fuerza la carga completa
    con Pillow antes de confiar en ella. Si algo falla (Pillow no
    instalado, datos corruptos, etc.), regresa None y el encabezado
    cae de vuelta al texto.
    """
    try:
        datos = base64.b64decode(LOGO_INE_PNG_BASE64)
        from PIL import Image as PILImage
        with PILImage.open(BytesIO(datos)) as pil_img:
            pil_img.load()
        return datos
    except Exception:
        return None


def F(**k):
    return Font(name="Arial", **k)


# ============================================================
# FECHAS Y ARCHIVOS
# ============================================================
def normaliza_fecha(texto):
    t = texto.replace("\u00a0", " ")
    t = re.sub(r"a\.\s*m\.", "AM", t, flags=re.IGNORECASE)
    t = re.sub(r"p\.\s*m\.", "PM", t, flags=re.IGNORECASE)
    return t.strip()


def parsea_fecha_cdmx(texto):
    t = normaliza_fecha(texto)
    dt_naive = None
    if HAVE_DATEUTIL:
        try:
            dt_naive = dtparser.parse(t, dayfirst=True)
        except (ValueError, OverflowError, TypeError):
            dt_naive = None
    if dt_naive is None:
        for fmt in ("%d/%m/%Y %I:%M:%S %p", "%d/%m/%Y %H:%M:%S", "%d/%m/%Y %I:%M %p", "%d/%m/%Y %H:%M"):
            try:
                dt_naive = datetime.strptime(t, fmt)
                break
            except ValueError:
                continue
    if dt_naive is None:
        return None
    return dt_naive.replace(tzinfo=TZ_CDMX)


def parsea_fecha_arg(texto):
    try:
        return datetime.strptime(texto.strip(), "%Y-%m-%d").date()
    except ValueError:
        raise argparse.ArgumentTypeError(
            f"Fecha inválida: '{texto}' (usa el formato AAAA-MM-DD, por ejemplo 2026-06-01)"
        )


def normaliza_texto(s):
    """Minúsculas y sin acentos, para comparar sin importar mayúsculas ni
    tildes (p. ej. 'Libélula' y 'libelula' deben coincidir)."""
    s = (s or "").strip().lower()
    s = unicodedata.normalize("NFKD", s)
    return "".join(c for c in s if not unicodedata.combining(c))


def coincide_fuzzy(direccion_correo, patron_norm, umbral=0.72):
    """
    True si `direccion_correo` "se parece" a `patron_norm` (ya normalizado):
    coincidencia exacta de subcadena, o similitud aproximada contra el
    nombre de usuario (parte antes de la @) y contra la dirección completa.
    """
    if not patron_norm:
        return False
    t = normaliza_texto(direccion_correo)
    if patron_norm in t:
        return True
    local = t.split("@", 1)[0] if "@" in t else t
    if difflib.SequenceMatcher(None, patron_norm, local).ratio() >= umbral:
        return True
    return difflib.SequenceMatcher(None, patron_norm, t).ratio() >= umbral


RE_CORREO = re.compile(r"[\w.+-]+@[\w-]+\.[\w.-]+")


def extrae_correo(texto):
    """
    Extrae la dirección de correo de un campo que puede venir como
    'usuario@dominio.com' (limpio) o como 'Nombre Apellido <usuario@dominio.com>'
    (con nombre para mostrar). Si no encuentra un patrón de correo, regresa
    el texto tal cual (recortado y en minúsculas) para no perder el dato.
    """
    t = (texto or "").strip()
    m = RE_CORREO.search(t)
    return m.group(0).lower() if m else t.lower()


RE_FECHA_ARCHIVO = re.compile(r"MessageTrace_(\d{8})")


def extrae_fecha_archivo(nombre):
    """Extrae la fecha (AAAAMMDD) del nombre de archivo, p. ej.
    'MessageTrace_20210827-070014.csv' -> date(2021, 8, 27). Regresa None
    si el nombre no trae esa fecha (para no arriesgarse a descartarlo mal)."""
    m = RE_FECHA_ARCHIVO.match(nombre)
    if not m:
        return None
    try:
        return datetime.strptime(m.group(1), "%Y%m%d").date()
    except ValueError:
        return None


def archivos_todos(desde=None, hasta=None):
    """
    Todos los MessageTrace_*.csv disponibles. Si se dan desde/hasta,
    descarta de entrada (por nombre de archivo, MessageTrace_AAAAMMDD-...)
    los que caen fuera del rango, sin necesidad de abrirlos. Un archivo
    cuyo nombre no trae fecha reconocible siempre se incluye.
    """
    candidatos = []
    for f in glob.glob(os.path.join(CSV_DIR, "MessageTrace_*.csv")):
        nombre = os.path.basename(f)
        if nombre in EXCLUDE_FILES:
            continue
        if desde is not None or hasta is not None:
            fecha_archivo = extrae_fecha_archivo(nombre)
            if fecha_archivo is not None:
                if desde is not None and fecha_archivo < desde:
                    continue
                if hasta is not None and fecha_archivo > hasta:
                    continue
        candidatos.append(f)
    return sorted(candidatos)


def lee_csv_robusto(path):
    for enc in ("utf-8", "cp1252"):
        try:
            with open(path, "r", encoding=enc, newline="") as fh:
                return list(csv.reader(fh))
        except UnicodeDecodeError:
            continue
    with open(path, "r", encoding="cp1252", errors="replace", newline="") as fh:
        return list(csv.reader(fh))


def formatea_duracion(segundos):
    segundos = max(0, int(round(segundos)))
    h, resto = divmod(segundos, 3600)
    m, s = divmod(resto, 60)
    if h:
        return f"{h}h {m:02d}m"
    if m:
        return f"{m}m {s:02d}s"
    return f"{s}s"


def imprime_barra_progreso(actual, total, tiempo_inicio, ancho=30):
    if total <= 0:
        return
    pct = actual / total
    llenado = int(ancho * pct)
    barra = "#" * llenado + "-" * (ancho - llenado)
    transcurrido = time.time() - tiempo_inicio
    if actual > 0:
        eta_txt = formatea_duracion((transcurrido / actual) * (total - actual)) if actual < total else "0s"
    else:
        eta_txt = "calculando..."
    sys.stdout.write(
        f"\r  [{barra}] {actual}/{total} archivos ({pct * 100:5.1f}%)"
        f" | transcurrido: {formatea_duracion(transcurrido)}"
        f" | restante (estimado): {eta_txt}   "
    )
    sys.stdout.flush()
    if actual >= total:
        sys.stdout.write("\n")
        sys.stdout.flush()


def _reune_resultados(archivos, funcion_por_archivo, progreso, mostrar_coincidencias, formatea_linea_match, hilos=8):
    """
    Motor común de recolección: procesa `archivos` en paralelo con
    `funcion_por_archivo(path) -> dict con matches`, deduplica por huella
    única y opcionalmente imprime cada coincidencia nueva en consola
    conforme se va encontrando. Comparte esta lógica busca_direccion,
    busca_todo y busca_patron para no repetirla tres veces.
    """
    total_archivos = len(archivos)
    resultados = []
    vistos = {}
    duplicados = 0
    ejemplos_duplicados = []
    MAX_EJEMPLOS_DUPLICADOS = 20
    archivos_omitidos_estructura = []
    archivos_procesados_ok = 0
    archivos_con_error = []
    tiempo_inicio_busqueda = time.time()

    completados = 0
    with concurrent.futures.ThreadPoolExecutor(max_workers=hilos) as executor:
        futuros = {executor.submit(funcion_por_archivo, path): path for path in archivos}
        for futuro in concurrent.futures.as_completed(futuros):
            completados += 1
            r = futuro.result()
            if r["error"]:
                archivos_con_error.append(f"{r['archivo']}: {r['error']}")
                if progreso:
                    imprime_barra_progreso(completados, total_archivos, tiempo_inicio_busqueda)
                continue
            if r["omitido_estructura"]:
                archivos_omitidos_estructura.append(r["omitido_estructura"])
                if progreso:
                    imprime_barra_progreso(completados, total_archivos, tiempo_inicio_busqueda)
                continue
            archivos_procesados_ok += 1

            nuevos = []
            for x in r["matches"]:
                s_low = x["sender"].strip().lower()
                r_low = x["recipient"].strip().lower()
                huella = (x["msgid"].strip(), x["fecha_original"], s_low, r_low, x["status"].strip().lower())
                if huella in vistos:
                    duplicados += 1
                    if len(ejemplos_duplicados) < MAX_EJEMPLOS_DUPLICADOS:
                        archivo_previo, subject_previo, size_previo = vistos[huella]
                        ejemplos_duplicados.append({
                            "msgid": x["msgid"].strip(),
                            "received": x["fecha_original"],
                            "sender": x["sender"],
                            "recipient": x["recipient"],
                            "status": x["status"],
                            "archivo_1": archivo_previo,
                            "subject_1": subject_previo,
                            "size_1": size_previo,
                            "archivo_2": x["_archivo_origen"],
                            "subject_2": x["subject"],
                            "size_2": x["size"],
                        })
                    continue
                vistos[huella] = (x["_archivo_origen"], x["subject"], x["size"])
                resultados.append(x)
                nuevos.append(x)

            if mostrar_coincidencias and nuevos:
                sys.stdout.write("\n")
                for x in nuevos:
                    print("  [ENCONTRADO] " + formatea_linea_match(x))
            if progreso:
                imprime_barra_progreso(completados, total_archivos, tiempo_inicio_busqueda)

    resultados.sort(key=lambda r: r["fecha_cdmx"])

    auditoria = {
        "archivos_candidatos": total_archivos,
        "archivos_procesados_ok": archivos_procesados_ok,
        "archivos_omitidos_estructura": archivos_omitidos_estructura,
        "archivos_con_error": archivos_con_error,
        "duplicados": duplicados,
        "ejemplos_duplicados": ejemplos_duplicados,
        "tiempo_busqueda_seg": time.time() - tiempo_inicio_busqueda,
    }
    return resultados, auditoria


def _linea_match_generica(x):
    fecha_txt = x["fecha_cdmx"].strftime("%d/%m/%Y %H:%M:%S")
    asunto = (x["subject"] or "(sin asunto)").strip()
    return f"{fecha_txt} | {x['sender']} -> {x['recipient']}  |  {asunto}"


# ============================================================
# MODO 1: ENVIADOS/RECIBIDOS A DOMINIOS EXTERNOS
# ============================================================
def _procesa_archivo_direccion(path, usuario, direccion, desde, hasta):
    nombre_archivo = os.path.basename(path)
    try:
        filas = lee_csv_robusto(path)
    except Exception as e:
        return {"archivo": nombre_archivo, "error": str(e), "omitido_estructura": None, "matches": []}
    if not filas:
        return {"archivo": nombre_archivo, "error": None, "omitido_estructura": None, "matches": []}
    if len(filas[0]) != 8:
        return {
            "archivo": nombre_archivo, "error": None,
            "omitido_estructura": f"{nombre_archivo} ({len(filas[0])} columnas, se esperaban 8)",
            "matches": [],
        }

    matches = []
    for fila in filas:
        if len(fila) != 8:
            continue
        received, sender, recipient, subject, status, fromip, size, msgid = fila[:8]
        if received.strip().lower() == "received":
            continue
        s_low, r_low = extrae_correo(sender), extrae_correo(recipient)

        if direccion == "enviados":
            if s_low != usuario:
                continue
            dominio_otro = r_low.rsplit("@", 1)[-1] if "@" in r_low else ""
        else:  # "recibidos"
            if r_low != usuario:
                continue
            dominio_otro = s_low.rsplit("@", 1)[-1] if "@" in s_low else ""
        if dominio_otro in DOMINIOS_INTERNOS:
            continue

        dt_cdmx = parsea_fecha_cdmx(received)
        if dt_cdmx is None:
            continue
        if desde is not None and dt_cdmx.date() < desde:
            continue
        if hasta is not None and dt_cdmx.date() > hasta:
            continue

        matches.append({
            "fecha_original": received.strip(), "fecha_cdmx": dt_cdmx, "fecha_utc": dt_cdmx.astimezone(TZ_UTC),
            "sender": sender, "recipient": recipient, "subject": subject, "status": status,
            "size": size, "msgid": msgid, "_archivo_origen": nombre_archivo,
        })
    return {"archivo": nombre_archivo, "error": None, "omitido_estructura": None, "matches": matches}


def busca_direccion(usuario, direccion, desde=None, hasta=None, progreso=True, hilos=8, mostrar_coincidencias=True):
    usuario = extrae_correo(usuario)
    archivos = archivos_todos(desde=desde, hasta=hasta)
    if progreso:
        rango_txt = (
            f"del {desde.strftime('%d/%m/%Y')} al {hasta.strftime('%d/%m/%Y')}"
            if (desde or hasta) else "TODO el histórico disponible"
        )
        print(f"Revisando {len(archivos)} archivo(s) — {rango_txt} — {direccion} por {usuario} (externos)...")

    def procesa(path):
        return _procesa_archivo_direccion(path, usuario, direccion, desde, hasta)
    return _reune_resultados(archivos, procesa, progreso, mostrar_coincidencias, _linea_match_generica, hilos=hilos)


# ============================================================
# MODO 2: TODO (enviados + recibidos, SIN excluir internos)
# ============================================================
def _procesa_archivo_todo(path, usuario, desde, hasta):
    nombre_archivo = os.path.basename(path)
    try:
        filas = lee_csv_robusto(path)
    except Exception as e:
        return {"archivo": nombre_archivo, "error": str(e), "omitido_estructura": None, "matches": []}
    if not filas:
        return {"archivo": nombre_archivo, "error": None, "omitido_estructura": None, "matches": []}
    if len(filas[0]) != 8:
        return {
            "archivo": nombre_archivo, "error": None,
            "omitido_estructura": f"{nombre_archivo} ({len(filas[0])} columnas, se esperaban 8)",
            "matches": [],
        }

    matches = []
    for fila in filas:
        if len(fila) != 8:
            continue
        received, sender, recipient, subject, status, fromip, size, msgid = fila[:8]
        if received.strip().lower() == "received":
            continue
        s_low, r_low = extrae_correo(sender), extrae_correo(recipient)
        if usuario != s_low and usuario != r_low:
            continue

        dt_cdmx = parsea_fecha_cdmx(received)
        if dt_cdmx is None:
            continue
        if desde is not None and dt_cdmx.date() < desde:
            continue
        if hasta is not None and dt_cdmx.date() > hasta:
            continue

        matches.append({
            "fecha_original": received.strip(), "fecha_cdmx": dt_cdmx, "fecha_utc": dt_cdmx.astimezone(TZ_UTC),
            "sender": sender, "recipient": recipient, "subject": subject, "status": status,
            "size": size, "msgid": msgid, "_archivo_origen": nombre_archivo,
        })
    return {"archivo": nombre_archivo, "error": None, "omitido_estructura": None, "matches": matches}


def busca_todo(usuario, desde=None, hasta=None, progreso=True, hilos=8, mostrar_coincidencias=True):
    usuario = extrae_correo(usuario)
    archivos = archivos_todos(desde=desde, hasta=hasta)
    if progreso:
        rango_txt = (
            f"del {desde.strftime('%d/%m/%Y')} al {hasta.strftime('%d/%m/%Y')}"
            if (desde or hasta) else "TODO el histórico disponible"
        )
        print(f"Revisando {len(archivos)} archivo(s) — {rango_txt} — TODO lo enviado/recibido por {usuario} (incluye internos)...")

    def procesa(path):
        return _procesa_archivo_todo(path, usuario, desde, hasta)
    return _reune_resultados(archivos, procesa, progreso, mostrar_coincidencias, _linea_match_generica, hilos=hilos)


# ============================================================
# MODO 3: BÚSQUEDA POR PATRÓN (cualquier remitente)
# ============================================================
def _procesa_archivo_patron(path, patron_norm, campo, desde, hasta):
    nombre_archivo = os.path.basename(path)
    try:
        filas = lee_csv_robusto(path)
    except Exception as e:
        return {"archivo": nombre_archivo, "error": str(e), "omitido_estructura": None, "matches": []}
    if not filas:
        return {"archivo": nombre_archivo, "error": None, "omitido_estructura": None, "matches": []}
    if len(filas[0]) != 8:
        return {
            "archivo": nombre_archivo, "error": None,
            "omitido_estructura": f"{nombre_archivo} ({len(filas[0])} columnas, se esperaban 8)",
            "matches": [],
        }

    matches = []
    for fila in filas:
        if len(fila) != 8:
            continue
        received, sender, recipient, subject, status, fromip, size, msgid = fila[:8]
        if received.strip().lower() == "received":
            continue

        objetivo = recipient if campo == "destinatario" else sender
        if not coincide_fuzzy(objetivo, patron_norm):
            continue

        dt_cdmx = parsea_fecha_cdmx(received)
        if dt_cdmx is None:
            continue
        if desde is not None and dt_cdmx.date() < desde:
            continue
        if hasta is not None and dt_cdmx.date() > hasta:
            continue

        matches.append({
            "fecha_original": received.strip(), "fecha_cdmx": dt_cdmx, "fecha_utc": dt_cdmx.astimezone(TZ_UTC),
            "sender": sender, "recipient": recipient, "subject": subject, "status": status,
            "size": size, "msgid": msgid, "_archivo_origen": nombre_archivo,
        })
    return {"archivo": nombre_archivo, "error": None, "omitido_estructura": None, "matches": matches}


def busca_patron(patron, campo="destinatario", desde=None, hasta=None, progreso=True, hilos=8, mostrar_coincidencias=True):
    patron_norm = normaliza_texto(patron)
    archivos = archivos_todos(desde=desde, hasta=hasta)
    if progreso:
        rango_txt = (
            f"del {desde.strftime('%d/%m/%Y')} al {hasta.strftime('%d/%m/%Y')}"
            if (desde or hasta) else "TODO el histórico disponible"
        )
        print(f"Revisando {len(archivos)} archivo(s) — {rango_txt} — buscando '{patron}' en {campo}...")

    def procesa(path):
        return _procesa_archivo_patron(path, patron_norm, campo, desde, hasta)
    return _reune_resultados(archivos, procesa, progreso, mostrar_coincidencias, _linea_match_generica, hilos=hilos)


# ============================================================
# EXCEL
# ============================================================
def oculta_no_usado(ws, last_col, last_row, col_buffer=40, row_buffer=300):
    for c in range(last_col + 1, last_col + 1 + col_buffer):
        ws.column_dimensions[get_column_letter(c)].hidden = True
    for rr in range(last_row + 1, last_row + 1 + row_buffer):
        ws.row_dimensions[rr].hidden = True


def _encabezado_institucional(ws, last_col):
    """Encabezado compartido por TODAS las hojas de TODOS los modos: logo
    del INE si la imagen embebida es válida, o el nombre en texto si no
    (así nunca se muestra el cuadro negro de una imagen corrupta)."""
    r = 1
    datos_logo = _logo_valido()
    if datos_logo:
        img = XLImage(BytesIO(datos_logo))
        ancho_orig, alto_orig = img.width, img.height
        ancho_px = 180
        alto_px = int(ancho_px * alto_orig / ancho_orig) if ancho_orig else 74
        img.width = ancho_px
        img.height = alto_px
        ws.row_dimensions[1].height = 32
        ws.row_dimensions[2].height = 32
        ws.add_image(img, "B1")
        r = 3
    else:
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=last_col)
        c = ws.cell(r, 2, "INSTITUTO NACIONAL ELECTORAL")
        c.font = F(bold=True, size=15, color=LILA_INST)
        r += 1

    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=last_col)
    ws.cell(r, 2, "Unidad Técnica de Servicios de Informática (UTSI)").font = F(size=11)
    r += 1
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=last_col)
    ws.cell(
        r, 2,
        "Departamento de Soporte Técnico y Administración de Servicios de"
        " Colaboración (DSTyASC)",
    ).font = F(size=10, italic=True)
    r += 2
    return r


def crea_hoja_detalle(ws, titulo_pestana, lista_datos):
    thin = Side(style="thin", color="BFBFBF")
    border_all = Border(left=thin, right=thin, top=thin, bottom=thin)
    border_bottom = Border(bottom=Side(style="thin", color="BFBFBF"))
    font_detalle = F(size=8)

    ws.sheet_view.showGridLines = False
    ancho_num = max(4, len(str(len(lista_datos))) + 2)
    widths = [ancho_num, 18, 18, 26, 26, 45, 14, 10]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    LAST_COL = 8
    r = _encabezado_institucional(ws, LAST_COL)

    c = ws.cell(r, 2, f"DETALLE DE MENSAJES {titulo_pestana.upper()}")
    c.font = F(bold=True, size=14, color=BLANCO)
    c.fill = PatternFill("solid", fgColor=LILA_INST)
    c.alignment = Alignment(horizontal="center")
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=LAST_COL)
    ws.row_dimensions[r].height = 24
    r += 2

    headers = ["#", "FECHA ORIGINAL", "FECHA (CDMX / UTC-6)", "REMITENTE", "DESTINATARIO", "ASUNTO", "ESTATUS", "TAMAÑO (KB)"]
    for j, h in enumerate(headers, start=1):
        c = ws.cell(r, j, h)
        c.font = F(bold=True, size=9, color=BLANCO)
        c.fill = PatternFill("solid", fgColor=LILA_OSCURO)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = border_all
    ws.row_dimensions[r].height = 22
    r += 1

    tabla_ini = r
    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")
    fill_zebra = PatternFill("solid", fgColor=GRIS_CLARO)

    for idx, x in enumerate(lista_datos, start=1):
        st_es = status_label(x["status"])
        try:
            kb = round(int(str(x["size"]).strip() or 0) / 1024, 1)
        except (ValueError, TypeError):
            kb = ""
        vals = [idx, x["fecha_original"], x["fecha_cdmx"].strftime("%d/%m/%Y %H:%M:%S"),
                x["sender"], x["recipient"], x["subject"], st_es, kb]
        zebra = idx % 2 == 0
        for j, v in enumerate(vals, start=1):
            c = ws.cell(r, j, v)
            c.font = font_detalle
            c.alignment = align_center if j in (1, 2, 3, 7, 8) else align_left
            c.border = border_bottom
            if zebra:
                c.fill = fill_zebra
        r += 1

    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_margins = PageMargins(left=0.4, right=0.4, top=0.5, bottom=0.5)
    ws.print_area = f"A1:{get_column_letter(LAST_COL)}{r - 1}"
    ws.print_title_rows = f"{tabla_ini - 1}:{tabla_ini - 1}"
    ws.freeze_panes = ws.cell(tabla_ini, 1)
    oculta_no_usado(ws, LAST_COL, r - 1)


def _bloque_resumen_status(ws, r, last_col, resultados, border_all):
    def bloque_titulo(row, texto):
        ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=last_col)
        c = ws.cell(row, 2, texto)
        c.font = F(bold=True, size=11, color=BLANCO)
        c.fill = PatternFill("solid", fgColor=LILA_OSCURO)
        c.alignment = Alignment(horizontal="left", indent=1, vertical="center")
        ws.row_dimensions[row].height = 20

    por_status = {}
    for x in resultados:
        st = status_label(x["status"])
        por_status[st] = por_status.get(st, 0) + 1
    if not por_status:
        return r

    ws.cell(r, 2, "Desglose por estatus").font = F(bold=True, size=10, color=LILA_OSCURO)
    r += 1
    hdr = ws.cell(r, 2, "Estatus")
    hdr.font = F(bold=True, size=9, color=BLANCO)
    hdr.fill = PatternFill("solid", fgColor=LILA_INST)
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=5)
    hdr2 = ws.cell(r, 6, "Cantidad")
    hdr2.font = F(bold=True, size=9, color=BLANCO)
    hdr2.fill = PatternFill("solid", fgColor=LILA_INST)
    ws.merge_cells(start_row=r, start_column=6, end_row=r, end_column=last_col)
    for col in range(2, last_col + 1):
        ws.cell(r, col).border = border_all
    r += 1
    for st, cnt in sorted(por_status.items(), key=lambda x: -x[1]):
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=5)
        ws.cell(r, 2, st).font = F(size=9)
        ws.merge_cells(start_row=r, start_column=6, end_row=r, end_column=last_col)
        c2 = ws.cell(r, 6, cnt)
        c2.font = F(size=9)
        c2.alignment = Alignment(horizontal="center")
        low = st.lower()
        fill = VERDE_OK if "entreg" in low else (ROJO_ERR if "error" in low else GRIS_CLARO)
        for col in range(2, last_col + 1):
            ws.cell(r, col).fill = PatternFill("solid", fgColor=fill)
            ws.cell(r, col).border = border_all
        r += 1
    r += 1

    raws_presentes = {}
    for x in resultados:
        raws_presentes[status_label(x["status"])] = x["status"]
    bloque_titulo(r, "NOMENCLATURA DE ESTATUS")
    r += 1
    for st_label in sorted(raws_presentes.keys()):
        desc = status_desc(raws_presentes[st_label])
        c1 = ws.cell(r, 2, st_label)
        c1.font = F(bold=True, size=9, color=LILA_OSCURO)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
        c1.alignment = Alignment(vertical="top")
        c2 = ws.cell(r, 4, desc)
        c2.font = F(size=9)
        ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=last_col)
        c2.alignment = Alignment(vertical="top", wrap_text=True)
        for col in range(2, last_col + 1):
            ws.cell(r, col).border = border_all
        ws.row_dimensions[r].height = 26
        r += 1
    return r


def _guarda_workbook(wb, out_path):
    try:
        wb.save(out_path)
    except PermissionError:
        base, ext = os.path.splitext(out_path)
        alterno = f"{base}_{datetime.now(TZ_CDMX).strftime('%H%M%S')}{ext}"
        print(f"[AVISO] No se pudo guardar en {out_path} (¿está abierto en Excel?). Se guardó como: {alterno}")
        wb.save(alterno)
        return alterno
    return out_path


def genera_excel_direccion(resultados, usuario, direccion, desde, hasta, out_path):
    """Modo 1: enviados o recibidos hacia/desde externos. 2 pestañas:
    Resumen + Detalle."""
    thin = Side(style="thin", color="BFBFBF")
    border_all = Border(left=thin, right=thin, top=thin, bottom=thin)

    wb = openpyxl.Workbook()
    ws1 = wb.active
    ws1.title = "Resumen"
    ws1.sheet_view.showGridLines = False

    LAST_COL = 7
    widths = [6, 20, 26, 26, 45, 14, 10]
    for i, w in enumerate(widths, start=1):
        ws1.column_dimensions[get_column_letter(i)].width = w

    r = _encabezado_institucional(ws1, LAST_COL)
    titulo_modo = "ENVIADOS" if direccion == "enviados" else "RECIBIDOS"
    c = ws1.cell(r, 2, f"REPORTE DE CORREOS {titulo_modo} A/DE DOMINIOS EXTERNOS")
    c.font = F(bold=True, size=14, color=BLANCO)
    c.fill = PatternFill("solid", fgColor=LILA_INST)
    c.alignment = Alignment(horizontal="center")
    ws1.merge_cells(start_row=r, start_column=2, end_row=r, end_column=LAST_COL)
    ws1.row_dimensions[r].height = 24
    r += 2

    def bloque_titulo(row, texto):
        ws1.merge_cells(start_row=row, start_column=2, end_row=row, end_column=LAST_COL)
        c = ws1.cell(row, 2, texto)
        c.font = F(bold=True, size=11, color=BLANCO)
        c.fill = PatternFill("solid", fgColor=LILA_OSCURO)
        c.alignment = Alignment(horizontal="left", indent=1, vertical="center")
        ws1.row_dimensions[row].height = 20

    bloque_titulo(r, "DATOS DE LA CONSULTA")
    r += 1
    if resultados:
        primer_fecha = min(x["fecha_cdmx"] for x in resultados).strftime("%d/%m/%Y")
        ultima_fecha = max(x["fecha_cdmx"] for x in resultados).strftime("%d/%m/%Y")
        rango_encontrado_txt = f"{primer_fecha} al {ultima_fecha} (horario CDMX)"
    else:
        rango_encontrado_txt = "Sin resultados en el rango/histórico revisado"
    rango_pedido_txt = (
        f"{desde.strftime('%d/%m/%Y') if desde else 'inicio del histórico'} al "
        f"{hasta.strftime('%d/%m/%Y') if hasta else 'hoy'}"
        if (desde or hasta) else "TODO el histórico disponible (sin filtro de fecha)"
    )
    etiqueta_modo = (
        "Correos ENVIADOS por el usuario hacia dominios externos"
        if direccion == "enviados" else "Correos RECIBIDOS por el usuario desde dominios externos"
    )
    datos = [
        ("Usuario consultado:", usuario),
        ("Modo de búsqueda:", etiqueta_modo),
        ("Dominios internos omitidos:", ", ".join(sorted(DOMINIOS_INTERNOS))),
        ("Rango solicitado:", rango_pedido_txt),
        ("Rango real encontrado:", rango_encontrado_txt),
        ("Fecha de generación:", datetime.now(TZ_CDMX).strftime("%d/%m/%Y %H:%M hrs (CDMX)")),
    ]
    for etiqueta, valor in datos:
        c1 = ws1.cell(r, 2, etiqueta)
        c1.font = F(bold=True, size=10)
        ws1.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
        ws1.merge_cells(start_row=r, start_column=4, end_row=r, end_column=LAST_COL)
        ws1.cell(r, 4, valor).font = F(size=10)
        r += 1
    r += 1

    bloque_titulo(r, "RESUMEN CONSOLIDADO")
    r += 1
    c1 = ws1.cell(r, 2, f"Total de correos {titulo_modo.lower()} a/de externos")
    c1.font = F(bold=True, size=10)
    ws1.merge_cells(start_row=r, start_column=2, end_row=r, end_column=5)
    ws1.merge_cells(start_row=r, start_column=6, end_row=r, end_column=LAST_COL)
    c2 = ws1.cell(r, 6, len(resultados))
    c2.font = F(size=13, bold=True, color=LILA_INST)
    c2.alignment = Alignment(horizontal="center")
    for col in range(2, LAST_COL + 1):
        ws1.cell(r, col).border = border_all
    ws1.row_dimensions[r].height = 20
    r += 2

    r = _bloque_resumen_status(ws1, r, LAST_COL, resultados, border_all)

    ws1.page_setup.orientation = "landscape"
    ws1.page_setup.fitToWidth = 1
    ws1.page_setup.fitToHeight = 0
    ws1.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws1.page_margins = PageMargins(left=0.4, right=0.4, top=0.5, bottom=0.5)
    ws1.print_area = f"A1:{get_column_letter(LAST_COL)}{r}"

    ws2 = wb.create_sheet(title=titulo_modo.capitalize())
    crea_hoja_detalle(ws2, titulo_modo.capitalize(), resultados)

    out_path = _guarda_workbook(wb, out_path)
    return len(resultados), out_path


def genera_excel_todo(resultados, usuario, desde, hasta, out_path):
    """Modo 2: enviados + recibidos, sin excluir internos. 3 pestañas:
    Reporte General, Enviados, Recibidos."""
    thin = Side(style="thin", color="BFBFBF")
    border_all = Border(left=thin, right=thin, top=thin, bottom=thin)

    wb = openpyxl.Workbook()
    enviados_list = [x for x in resultados if x["sender"].strip().lower() == usuario.lower()]
    recibidos_list = [x for x in resultados if x["recipient"].strip().lower() == usuario.lower()]

    ws1 = wb.active
    ws1.title = "Reporte General"
    ws1.sheet_view.showGridLines = False

    LAST_COL = 7
    widths = [6, 20, 26, 26, 45, 14, 10]
    for i, w in enumerate(widths, start=1):
        ws1.column_dimensions[get_column_letter(i)].width = w

    r = _encabezado_institucional(ws1, LAST_COL)
    c = ws1.cell(r, 2, "REPORTE DE BITÁCORA DE CORREO ELECTRÓNICO (TODO, INCLUYE INTERNOS)")
    c.font = F(bold=True, size=14, color=BLANCO)
    c.fill = PatternFill("solid", fgColor=LILA_INST)
    c.alignment = Alignment(horizontal="center")
    ws1.merge_cells(start_row=r, start_column=2, end_row=r, end_column=LAST_COL)
    ws1.row_dimensions[r].height = 24
    r += 2

    def bloque_titulo(row, texto):
        ws1.merge_cells(start_row=row, start_column=2, end_row=row, end_column=LAST_COL)
        c = ws1.cell(row, 2, texto)
        c.font = F(bold=True, size=11, color=BLANCO)
        c.fill = PatternFill("solid", fgColor=LILA_OSCURO)
        c.alignment = Alignment(horizontal="left", indent=1, vertical="center")
        ws1.row_dimensions[row].height = 20

    bloque_titulo(r, "DATOS DE LA CONSULTA")
    r += 1
    if resultados:
        primer_fecha = min(x["fecha_cdmx"] for x in resultados).strftime("%d/%m/%Y")
        ultima_fecha = max(x["fecha_cdmx"] for x in resultados).strftime("%d/%m/%Y")
        rango_encontrado_txt = f"{primer_fecha} al {ultima_fecha} (horario CDMX)"
    else:
        rango_encontrado_txt = "Sin resultados en el rango/histórico revisado"
    rango_pedido_txt = (
        f"{desde.strftime('%d/%m/%Y') if desde else 'inicio del histórico'} al "
        f"{hasta.strftime('%d/%m/%Y') if hasta else 'hoy'}"
        if (desde or hasta) else "TODO el histórico disponible (sin filtro de fecha)"
    )
    datos = [
        ("Usuario consultado:", usuario),
        ("Modo de búsqueda:", "TODO lo enviado y recibido (incluye correos internos↔internos)"),
        ("Rango solicitado:", rango_pedido_txt),
        ("Rango real encontrado:", rango_encontrado_txt),
        ("Fecha de generación:", datetime.now(TZ_CDMX).strftime("%d/%m/%Y %H:%M hrs (CDMX)")),
    ]
    for etiqueta, valor in datos:
        c1 = ws1.cell(r, 2, etiqueta)
        c1.font = F(bold=True, size=10)
        ws1.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
        ws1.merge_cells(start_row=r, start_column=4, end_row=r, end_column=LAST_COL)
        ws1.cell(r, 4, valor).font = F(size=10)
        r += 1
    r += 1

    total = len(resultados)
    bloque_titulo(r, "RESUMEN CONSOLIDADO")
    r += 1
    for etiqueta, valor in [
        ("Total de mensajes", total),
        ("Enviados por el usuario", len(enviados_list)),
        ("Recibidos por el usuario", len(recibidos_list)),
    ]:
        c1 = ws1.cell(r, 2, etiqueta)
        c1.font = F(bold=True, size=10)
        ws1.merge_cells(start_row=r, start_column=2, end_row=r, end_column=5)
        ws1.merge_cells(start_row=r, start_column=6, end_row=r, end_column=LAST_COL)
        c2 = ws1.cell(r, 6, valor)
        c2.font = F(size=13, bold=True, color=LILA_INST)
        c2.alignment = Alignment(horizontal="center")
        for col in range(2, LAST_COL + 1):
            ws1.cell(r, col).border = border_all
        ws1.row_dimensions[r].height = 20
        r += 1
    r += 1

    # Desglose por día (rango corto) o por mes (rango largo)
    por_dia = {}
    for x in enviados_list:
        d = x["fecha_cdmx"].date()
        por_dia.setdefault(d, {"enviados": 0, "recibidos": 0})
        por_dia[d]["enviados"] += 1
    for x in recibidos_list:
        d = x["fecha_cdmx"].date()
        por_dia.setdefault(d, {"enviados": 0, "recibidos": 0})
        por_dia[d]["recibidos"] += 1

    dias_rango = []
    if resultados:
        _d = min(x["fecha_cdmx"] for x in resultados).date()
        _hasta_d = max(x["fecha_cdmx"] for x in resultados).date()
        while _d <= _hasta_d:
            dias_rango.append(_d)
            _d += timedelta(days=1)

    UMBRAL_DIAS_PARA_AGRUPAR_POR_MES = 31
    if len(dias_rango) <= UMBRAL_DIAS_PARA_AGRUPAR_POR_MES:
        titulo_resumen = "RESUMEN POR DÍA"
        etiqueta_col = "Fecha"
        dow_es = ["Lun", "Mar", "Mié", "Jue", "Vie", "Sáb", "Dom"]
        periodos = []
        for dia in dias_rango:
            datos_dia = por_dia.get(dia, {"enviados": 0, "recibidos": 0})
            label = f"{dia.strftime('%d/%m/%Y')} ({dow_es[dia.weekday()]})"
            periodos.append((label, datos_dia["enviados"], datos_dia["recibidos"]))
    else:
        titulo_resumen = "RESUMEN POR MES"
        etiqueta_col = "Mes"
        MESES_ES = ["", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
                    "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
        por_mes = {}
        for x in enviados_list:
            key = (x["fecha_cdmx"].year, x["fecha_cdmx"].month)
            por_mes.setdefault(key, {"enviados": 0, "recibidos": 0})
            por_mes[key]["enviados"] += 1
        for x in recibidos_list:
            key = (x["fecha_cdmx"].year, x["fecha_cdmx"].month)
            por_mes.setdefault(key, {"enviados": 0, "recibidos": 0})
            por_mes[key]["recibidos"] += 1
        meses_rango = []
        if resultados:
            _min_d = min(x["fecha_cdmx"] for x in resultados).date()
            _max_d = max(x["fecha_cdmx"] for x in resultados).date()
            y, m = _min_d.year, _min_d.month
            while (y, m) <= (_max_d.year, _max_d.month):
                meses_rango.append((y, m))
                m += 1
                if m == 13:
                    m = 1
                    y += 1
        periodos = []
        for (y, m) in meses_rango:
            datos_mes = por_mes.get((y, m), {"enviados": 0, "recibidos": 0})
            label = f"{MESES_ES[m]} {y}"
            periodos.append((label, datos_mes["enviados"], datos_mes["recibidos"]))

    bloque_titulo(r, titulo_resumen)
    r += 1
    ws1.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
    c = ws1.cell(r, 2, etiqueta_col)
    c.font = F(bold=True, size=9, color=BLANCO)
    c.fill = PatternFill("solid", fgColor=LILA_INST)
    c.alignment = Alignment(horizontal="center")
    for j, lbl in zip((4, 5), ("Enviados", "Recibidos")):
        c = ws1.cell(r, j, lbl)
        c.font = F(bold=True, size=9, color=BLANCO)
        c.fill = PatternFill("solid", fgColor=LILA_INST)
        c.alignment = Alignment(horizontal="center")
    ws1.merge_cells(start_row=r, start_column=6, end_row=r, end_column=7)
    c = ws1.cell(r, 6, "Total")
    c.font = F(bold=True, size=9, color=BLANCO)
    c.fill = PatternFill("solid", fgColor=LILA_INST)
    c.alignment = Alignment(horizontal="center")
    for col in range(2, LAST_COL + 1):
        ws1.cell(r, col).border = border_all
    r += 1
    for label, env_d, rec_d in periodos:
        tot_d = env_d + rec_d
        sin_actividad = tot_d == 0
        color_txt = "999999" if sin_actividad else "000000"
        ws1.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
        c = ws1.cell(r, 2, label)
        c.font = F(size=9, color=color_txt)
        c.alignment = Alignment(horizontal="left", indent=1)
        c = ws1.cell(r, 4, env_d)
        c.font = F(size=9, color=color_txt)
        c.alignment = Alignment(horizontal="center")
        c = ws1.cell(r, 5, rec_d)
        c.font = F(size=9, color=color_txt)
        c.alignment = Alignment(horizontal="center")
        ws1.merge_cells(start_row=r, start_column=6, end_row=r, end_column=7)
        c = ws1.cell(r, 6, tot_d)
        c.font = F(size=9, bold=not sin_actividad, color=color_txt if sin_actividad else LILA_INST)
        c.alignment = Alignment(horizontal="center")
        for col in range(2, LAST_COL + 1):
            ws1.cell(r, col).border = border_all
            if sin_actividad:
                ws1.cell(r, col).fill = PatternFill("solid", fgColor=GRIS_CLARO)
        r += 1
    r += 1

    r = _bloque_resumen_status(ws1, r, LAST_COL, resultados, border_all)

    ws1.page_setup.orientation = "landscape"
    ws1.page_setup.fitToWidth = 1
    ws1.page_setup.fitToHeight = 0
    ws1.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws1.page_margins = PageMargins(left=0.4, right=0.4, top=0.5, bottom=0.5)
    ws1.print_area = f"A1:{get_column_letter(LAST_COL)}{r}"

    ws2 = wb.create_sheet(title="Enviados")
    crea_hoja_detalle(ws2, "Enviados", enviados_list)
    ws3 = wb.create_sheet(title="Recibidos")
    crea_hoja_detalle(ws3, "Recibidos", recibidos_list)

    out_path = _guarda_workbook(wb, out_path)
    return len(resultados), out_path


def genera_excel_patron(resultados, patron, campo, desde, hasta, out_path):
    """Modo 3: búsqueda por patrón difuso. 2 pestañas: Resumen +
    Coincidencias."""
    thin = Side(style="thin", color="BFBFBF")
    border_all = Border(left=thin, right=thin, top=thin, bottom=thin)

    wb = openpyxl.Workbook()
    ws1 = wb.active
    ws1.title = "Resumen"
    ws1.sheet_view.showGridLines = False

    LAST_COL = 7
    widths = [6, 20, 26, 26, 45, 14, 10]
    for i, w in enumerate(widths, start=1):
        ws1.column_dimensions[get_column_letter(i)].width = w

    r = _encabezado_institucional(ws1, LAST_COL)
    c = ws1.cell(r, 2, "REPORTE DE BÚSQUEDA POR PATRÓN EN BITÁCORA DE CORREO")
    c.font = F(bold=True, size=14, color=BLANCO)
    c.fill = PatternFill("solid", fgColor=LILA_INST)
    c.alignment = Alignment(horizontal="center")
    ws1.merge_cells(start_row=r, start_column=2, end_row=r, end_column=LAST_COL)
    ws1.row_dimensions[r].height = 24
    r += 2

    def bloque_titulo(row, texto):
        ws1.merge_cells(start_row=row, start_column=2, end_row=row, end_column=LAST_COL)
        c = ws1.cell(row, 2, texto)
        c.font = F(bold=True, size=11, color=BLANCO)
        c.fill = PatternFill("solid", fgColor=LILA_OSCURO)
        c.alignment = Alignment(horizontal="left", indent=1, vertical="center")
        ws1.row_dimensions[row].height = 20

    bloque_titulo(r, "DATOS DE LA CONSULTA")
    r += 1
    if resultados:
        primer_fecha = min(x["fecha_cdmx"] for x in resultados).strftime("%d/%m/%Y")
        ultima_fecha = max(x["fecha_cdmx"] for x in resultados).strftime("%d/%m/%Y")
        rango_encontrado_txt = f"{primer_fecha} al {ultima_fecha} (horario CDMX)"
    else:
        rango_encontrado_txt = "Sin resultados en el rango/histórico revisado"
    rango_pedido_txt = (
        f"{desde.strftime('%d/%m/%Y') if desde else 'inicio del histórico'} al "
        f"{hasta.strftime('%d/%m/%Y') if hasta else 'hoy'}"
        if (desde or hasta) else "TODO el histórico disponible (sin filtro de fecha)"
    )
    campo_txt = (
        "Destinatario (Remitente = cualquiera)" if campo == "destinatario"
        else "Remitente (Destinatario = cualquiera)"
    )
    datos = [
        ("Patrón buscado:", patron),
        ("Campo comparado:", campo_txt),
        ("Rango solicitado:", rango_pedido_txt),
        ("Rango real encontrado:", rango_encontrado_txt),
        ("Fecha de generación:", datetime.now(TZ_CDMX).strftime("%d/%m/%Y %H:%M hrs (CDMX)")),
        ("Nota:", "La coincidencia es aproximada (tolera acentos, mayúsculas y pequeñas variantes de escritura)."),
    ]
    for etiqueta, valor in datos:
        c1 = ws1.cell(r, 2, etiqueta)
        c1.font = F(bold=True, size=10)
        ws1.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
        ws1.merge_cells(start_row=r, start_column=4, end_row=r, end_column=LAST_COL)
        ws1.cell(r, 4, valor).font = F(size=10)
        r += 1
    r += 1

    bloque_titulo(r, "RESUMEN CONSOLIDADO")
    r += 1
    c1 = ws1.cell(r, 2, "Total de coincidencias")
    c1.font = F(bold=True, size=10)
    ws1.merge_cells(start_row=r, start_column=2, end_row=r, end_column=5)
    ws1.merge_cells(start_row=r, start_column=6, end_row=r, end_column=LAST_COL)
    c2 = ws1.cell(r, 6, len(resultados))
    c2.font = F(size=13, bold=True, color=LILA_INST)
    c2.alignment = Alignment(horizontal="center")
    for col in range(2, LAST_COL + 1):
        ws1.cell(r, col).border = border_all
    ws1.row_dimensions[r].height = 20
    r += 2

    r = _bloque_resumen_status(ws1, r, LAST_COL, resultados, border_all)

    ws1.page_setup.orientation = "landscape"
    ws1.page_setup.fitToWidth = 1
    ws1.page_setup.fitToHeight = 0
    ws1.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws1.page_margins = PageMargins(left=0.4, right=0.4, top=0.5, bottom=0.5)
    ws1.print_area = f"A1:{get_column_letter(LAST_COL)}{r}"

    ws2 = wb.create_sheet(title="Coincidencias")
    crea_hoja_detalle(ws2, "coincidentes", resultados)

    out_path = _guarda_workbook(wb, out_path)
    return len(resultados), out_path


def escribe_log_auditoria(etiqueta, modo_desc, auditoria, total_encontrados, xlsx_path):
    """
    Escribe (agrega) el detalle técnico de la búsqueda a un log propio en
    OUT_DIR/Logs/<etiqueta>.log. El Excel/PDF entregable solo trae la
    información pedida; este log es el rastro auditable de cómo se hizo la
    búsqueda: qué se revisó, qué se omitió y por qué.
    """
    log_dir = os.path.join(OUT_DIR, "Logs")
    os.makedirs(log_dir, exist_ok=True)
    safe = re.sub(r"[^a-zA-Z0-9]", "_", etiqueta.strip().lower())
    log_path = os.path.join(log_dir, f"{safe}.log")

    omitidos = auditoria["archivos_omitidos_estructura"]
    con_error = auditoria.get("archivos_con_error") or []

    lineas = []
    lineas.append("=" * 78)
    lineas.append(f"Auditoría de búsqueda — {datetime.now(TZ_CDMX).strftime('%d/%m/%Y %H:%M:%S hrs (CDMX)')}")
    lineas.append(f"Consulta           : {etiqueta}")
    lineas.append(f"Modo de búsqueda   : {modo_desc}")
    lineas.append("-" * 78)
    lineas.append(f"Archivos fuente revisados : {auditoria['archivos_procesados_ok']} de {auditoria['archivos_candidatos']} candidatos")
    lineas.append(f"Archivos con estructura no estándar (omitidos) : {len(omitidos) if omitidos else 'Ninguno'}")
    for nombre in omitidos:
        lineas.append(f"    - {nombre}")
    lineas.append(f"Archivos con error de lectura : {len(con_error) if con_error else 'Ninguno'}")
    for nombre in con_error:
        lineas.append(f"    - {nombre}")
    lineas.append(f"Registros duplicados detectados y omitidos : {auditoria['duplicados']}")
    lineas.append(f"Tiempo de búsqueda : {formatea_duracion(auditoria['tiempo_busqueda_seg'])}")
    lineas.append("-" * 78)
    lineas.append(f"Mensajes únicos encontrados : {total_encontrados}")
    lineas.append(f"Reporte generado : {xlsx_path}")
    lineas.append("=" * 78)
    lineas.append("")

    with open(log_path, "a", encoding="utf-8") as fh:
        fh.write("\n".join(lineas) + "\n")
    return log_path


def exporta_pdf(xlsx_path, pdf_path):
    try:
        import win32com.client
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        wb = excel.Workbooks.Open(xlsx_path)
        wb.Sheets.Select()
        wb.ActiveSheet.ExportAsFixedFormat(0, pdf_path)
        wb.Close(False)
        excel.Quit()
        return True
    except Exception as e:
        print(f"[AVISO] No se pudo generar el PDF automáticamente: {e}")
        print("        Abre el .xlsx en Excel y usa 'Guardar como > PDF' manualmente.")
        return False


# ============================================================
# MAIN
# ============================================================
def main():
    global CSV_DIR
    tiempo_inicio_total = time.time()
    ap = argparse.ArgumentParser(
        description=(
            "Consulta bitácoras de correo en 3 modos: enviados/recibidos a"
            " externos, TODO (incluye internos), o búsqueda por PATRÓN"
            " (cualquier remitente)."
        )
    )
    ap.add_argument("--usuario", help="Correo del usuario a buscar (modos 'direccion' y 'todo')")
    ap.add_argument("--direccion", choices=["enviados", "recibidos"], default=None,
                     help="Modo 1: enviados/recibidos hacia/desde dominios EXTERNOS.")
    ap.add_argument("--todo", action="store_true",
                     help="Modo 2: TODO lo enviado y recibido por el usuario, SIN excluir dominios internos.")
    ap.add_argument("--patron", default=None,
                     help="Modo 3: busca cualquier correo (de cualquier remitente) parecido a esta palabra.")
    ap.add_argument("--campo", choices=["destinatario", "remitente"], default="destinatario",
                     help="Solo con --patron: en qué campo buscar el parecido (default: destinatario).")
    ap.add_argument("--desde", type=parsea_fecha_arg, default=None, help="Fecha inicial AAAA-MM-DD (opcional, todos los modos).")
    ap.add_argument("--hasta", type=parsea_fecha_arg, default=None, help="Fecha final AAAA-MM-DD (opcional, todos los modos).")
    ap.add_argument("--csv-dir", dest="csv_dir", default=None, help=f"Carpeta con los MessageTrace_*.csv (default: {CSV_DIR}).")
    ap.add_argument("--hilos", type=int, default=8, help="Archivos a procesar en paralelo (default 8).")
    ap.add_argument("--silencioso", action="store_true",
                     help="No imprimir cada coincidencia en consola conforme se va encontrando (solo el resumen final).")
    args = ap.parse_args()

    if args.csv_dir:
        CSV_DIR = args.csv_dir
    print(f"Leyendo bitácoras desde: {CSV_DIR}")

    if args.desde and args.hasta and args.desde > args.hasta:
        print("[ERROR] --desde no puede ser posterior a --hasta.")
        sys.exit(1)

    mostrar = not args.silencioso

    # -------------------- MODO 3: PATRÓN --------------------
    if args.patron:
        resultados, auditoria = busca_patron(
            args.patron, campo=args.campo, desde=args.desde, hasta=args.hasta,
            hilos=args.hilos, mostrar_coincidencias=mostrar,
        )
        print(f"Encontrados: {len(resultados)} mensaje(s) único(s).")
        print(f"Archivos revisados: {auditoria['archivos_procesados_ok']} de {auditoria['archivos_candidatos']} candidatos.")
        if auditoria["archivos_omitidos_estructura"]:
            print(f"[ATENCIÓN] {len(auditoria['archivos_omitidos_estructura'])} archivo(s) con estructura distinta a la esperada:")
            for nombre in auditoria["archivos_omitidos_estructura"]:
                print(f"    - {nombre}")
        if auditoria.get("archivos_con_error"):
            print(f"[ATENCIÓN] {len(auditoria['archivos_con_error'])} archivo(s) con error de lectura:")
            for nombre in auditoria["archivos_con_error"]:
                print(f"    - {nombre}")
        if auditoria["duplicados"]:
            print(f"Se detectaron y omitieron {auditoria['duplicados']} registro(s) duplicado(s) entre los archivos.")
        print(f"Tiempo de búsqueda: {formatea_duracion(auditoria['tiempo_busqueda_seg'])}")

        etiqueta_log = f"patron_{args.patron}"
        modo_desc = f"Búsqueda por patrón '{args.patron}' en {args.campo} (cualquier remitente)"
        if not resultados:
            print("No se generará reporte: no hubo coincidencias para ese patrón.")
            log_path = escribe_log_auditoria(etiqueta_log, modo_desc, auditoria, 0, "N/A (sin resultados)")
            print(f"Auditoría agregada al log: {log_path}")
            print(f"Tiempo total: {formatea_duracion(time.time() - tiempo_inicio_total)}")
            sys.exit(0)

        safe_patron = re.sub(r"[^a-zA-Z0-9]", "_", args.patron)
        sufijo_fechas = (
            f"_{args.desde.strftime('%Y%m%d') if args.desde else 'inicio'}-{args.hasta.strftime('%Y%m%d') if args.hasta else 'hoy'}"
            if (args.desde or args.hasta) else "_TODO"
        )
        xlsx_path = os.path.join(OUT_DIR, f"CorreosPatron_{args.campo}_{safe_patron}{sufijo_fechas}.xlsx")
        total, xlsx_path = genera_excel_patron(resultados, args.patron, args.campo, args.desde, args.hasta, xlsx_path)
        print(f"Excel generado: {xlsx_path}")
        log_path = escribe_log_auditoria(etiqueta_log, modo_desc, auditoria, len(resultados), xlsx_path)
        print(f"Auditoría agregada al log: {log_path}")
        pdf_path = os.path.splitext(xlsx_path)[0] + ".pdf"
        if exporta_pdf(xlsx_path, pdf_path):
            print(f"PDF generado: {pdf_path}")
        print(f"Tiempo total: {formatea_duracion(time.time() - tiempo_inicio_total)}")
        return

    # -------------------- MODOS 1 y 2: por usuario --------------------
    usuario = args.usuario or input("Usuario (correo): ").strip()

    modo = "todo" if args.todo else args.direccion
    if not modo:
        while modo not in ("enviados", "recibidos", "todo"):
            modo = input(
                "¿Buscar ENVIADOS a externos, RECIBIDOS de externos, o TODO"
                " (enviados+recibidos, incluye internos)? (enviados/recibidos/todo): "
            ).strip().lower()

    if modo == "todo":
        resultados, auditoria = busca_todo(
            usuario, desde=args.desde, hasta=args.hasta, hilos=args.hilos, mostrar_coincidencias=mostrar,
        )
        modo_desc = "TODO lo enviado y recibido (incluye correos internos↔internos)"
    else:
        resultados, auditoria = busca_direccion(
            usuario, modo, desde=args.desde, hasta=args.hasta, hilos=args.hilos, mostrar_coincidencias=mostrar,
        )
        modo_desc = (
            "Correos ENVIADOS por el usuario hacia dominios externos" if modo == "enviados"
            else "Correos RECIBIDOS por el usuario desde dominios externos"
        )

    print(f"Encontrados: {len(resultados)} mensaje(s) único(s).")
    print(f"Archivos revisados: {auditoria['archivos_procesados_ok']} de {auditoria['archivos_candidatos']} candidatos.")
    if auditoria["archivos_omitidos_estructura"]:
        print(f"[ATENCIÓN] {len(auditoria['archivos_omitidos_estructura'])} archivo(s) con estructura distinta a la esperada:")
        for nombre in auditoria["archivos_omitidos_estructura"]:
            print(f"    - {nombre}")
    if auditoria.get("archivos_con_error"):
        print(f"[ATENCIÓN] {len(auditoria['archivos_con_error'])} archivo(s) con error de lectura:")
        for nombre in auditoria["archivos_con_error"]:
            print(f"    - {nombre}")
    if auditoria["duplicados"]:
        print(f"Se detectaron y omitieron {auditoria['duplicados']} registro(s) duplicado(s) entre los archivos.")
    print(f"Tiempo de búsqueda: {formatea_duracion(auditoria['tiempo_busqueda_seg'])}")

    if not resultados:
        print("No se generará reporte: no hubo resultados para ese usuario/modo/rango.")
        log_path = escribe_log_auditoria(usuario, modo_desc, auditoria, 0, "N/A (sin resultados)")
        print(f"Auditoría agregada al log: {log_path}")
        print(f"Tiempo total: {formatea_duracion(time.time() - tiempo_inicio_total)}")
        sys.exit(0)

    safe_user = re.sub(r"[^a-zA-Z0-9]", "_", usuario)
    sufijo_fechas = (
        f"_{args.desde.strftime('%Y%m%d') if args.desde else 'inicio'}-{args.hasta.strftime('%Y%m%d') if args.hasta else 'hoy'}"
        if (args.desde or args.hasta) else "_TODO"
    )
    if modo == "todo":
        xlsx_path = os.path.join(OUT_DIR, f"BitacoraCompleta_{safe_user}{sufijo_fechas}.xlsx")
        total, xlsx_path = genera_excel_todo(resultados, usuario, args.desde, args.hasta, xlsx_path)
    else:
        xlsx_path = os.path.join(OUT_DIR, f"CorreosExternos_{modo}_{safe_user}{sufijo_fechas}.xlsx")
        total, xlsx_path = genera_excel_direccion(resultados, usuario, modo, args.desde, args.hasta, xlsx_path)
    print(f"Excel generado: {xlsx_path}")

    log_path = escribe_log_auditoria(usuario, modo_desc, auditoria, len(resultados), xlsx_path)
    print(f"Auditoría agregada al log: {log_path}")

    pdf_path = os.path.splitext(xlsx_path)[0] + ".pdf"
    if exporta_pdf(xlsx_path, pdf_path):
        print(f"PDF generado: {pdf_path}")

    print(f"Tiempo total: {formatea_duracion(time.time() - tiempo_inicio_total)}")


if __name__ == "__main__":
    main()
