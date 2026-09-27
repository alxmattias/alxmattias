#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Buscar-Bitacoras.py
====================
Busca bitácoras de correo (Exchange Message Trace) por usuario y rango de
fechas, directamente sobre los CSV que están en esta misma carpeta, y genera
un reporte Excel con 3 pestañas (Reporte General, Enviados y Recibidos) + PDF
con el formato institucional INE.

USO:
    python Buscar-Bitacoras.py --usuario correo@ine.mx --desde 2026-06-01 --hasta 2026-06-15
    python Buscar-Bitacoras.py                     (modo interactivo, sin argumentos)

Requisitos (una sola vez):
    pip install openpyxl python-dateutil tzdata pywin32
"""

import argparse
import base64
import csv
import glob
import os
import re
import sys
import time
from datetime import datetime, timedelta
from io import BytesIO
from zoneinfo import ZoneInfo

# Sube el límite de tamaño de campo del módulo csv
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

# Ruta de LECTURA de los CSV de bitácoras: el servidor origen (VM-ADMON-WIN),
# vía red. Fija, no se mueve de aquí aunque el script se mueva de carpeta.
# Se puede sobreescribir en cada corrida con --csv-dir, por ejemplo para
# apuntar a la copia local E:\Bitacoras si el servidor remoto no está
# disponible.
CSV_DIR = r"\\VM-ADMON-WIN\E$\Bitacoras"

# El Excel/PDF resultante y este mismo script viven en su propia carpeta
# (por ejemplo E:\Bitacoras\Excel-Bitacoras), separada de los CSV fuente.
OUT_DIR = SCRIPT_DIR

TZ_CDMX = ZoneInfo("America/Mexico_City")
TZ_UTC = ZoneInfo("UTC")

EXCLUDE_FILES = {"fechas.csv", "usuarios.csv", "refechas.csv"}

STATUS_INFO = {
    "delivered": (
        "Entregado",
        "El mensaje llegó correctamente al buzón del destinatario.",
    ),
    "failed": (
        "Error",
        "El mensaje no pudo entregarse; fue rechazado o rebotado.",
    ),
    "pending": ("Pendiente", "El mensaje todavía está en proceso de entrega."),
    "resolved": (
        "Resuelto",
        "Notificación de un sistema de monitoreo que ya fue atendida.",
    ),
    "quarantined": (
        "En cuarentena",
        "El sistema de seguridad detuvo el mensaje por sospecha de spam o virus.",
    ),
    "filteredasspam": (
        "Filtrado como spam",
        "Se identificó como correo no deseado y no llegó a la bandeja principal.",
    ),
    "gettingstatus": (
        "Consultando estatus",
        "El sistema todavía está verificando qué pasó con el mensaje.",
    ),
    "expanded": (
        "Expandido",
        "El mensaje se envió a una lista o grupo de correo y se distribuyó a sus integrantes.",
    ),
    "recalled": (
        "Recall solicitado",
        "El remitente intentó recuperar (cancelar) este mensaje después de haberlo enviado. "
        "Si el destinatario ya lo había abierto, es posible que de todas formas lo haya visto.",
    ),
    "removedbymessagerecall": (
        "Retirado por recall",
        "El mensaje fue eliminado del buzón del destinatario porque el remitente lo recuperó a tiempo.",
    ),
    "poison": (
        "Mensaje dañado",
        "El sistema detectó un problema grave en el mensaje y lo detuvo por seguridad.",
    ),
    "none": (
        "Sin estatus",
        "No hay información de estatus registrada para este evento.",
    ),
}


def status_label(status_raw):
    info = STATUS_INFO.get((status_raw or "").strip().lower())
    return info[0] if info else (status_raw or "Sin estatus")


def status_desc(status_raw):
    info = STATUS_INFO.get((status_raw or "").strip().lower())
    return (
        info[1]
        if info
        else "Estatus no documentado; consulte con el área de soporte si tiene"
        " dudas."
    )


# Paleta institucional INE
LILA_INST = "674092"
LILA_OSCURO = "49276F"
LILA = "9680B4"
LILA_CLARO = "D1BDEF"
AZUL_INST = "383B7E"
GRIS_INE = "A6A8AA"
GRIS_CLARO = "F2F2F2"
BLANCO = "FFFFFF"
VERDE_OK = "E2EFDA"
ROJO_ERR = "FCE4E4"


# ============================================================
# PARSEO DE FECHAS Y BÚSQUEDA
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
        for fmt in (
            "%d/%m/%Y %I:%M:%S %p",
            "%d/%m/%Y %H:%M:%S",
            "%d/%m/%Y %I:%M %p",
            "%d/%m/%Y %H:%M",
        ):
            try:
                dt_naive = datetime.strptime(t, fmt)
                break
            except ValueError:
                continue

    if dt_naive is None:
        return None

    return dt_naive.replace(tzinfo=TZ_CDMX)


def archivos_candidatos(desde, hasta):
    inicio = desde - timedelta(days=1)
    fin = hasta + timedelta(days=1)
    candidatos = []
    for f in glob.glob(os.path.join(CSV_DIR, "MessageTrace_*.csv")):
        nombre = os.path.basename(f)
        if nombre in EXCLUDE_FILES:
            continue
        m = re.search(r"(\d{8})", nombre)
        if not m:
            continue
        try:
            d = datetime.strptime(m.group(1), "%Y%m%d").date()
        except ValueError:
            continue
        if inicio <= d <= fin:
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


def dias_sin_cobertura(desde, hasta, archivos):
    """
    Días dentro del rango pedido que NO tienen ningún archivo fuente que
    pudiera cubrirlos (ni con el nombre exacto del día, ni con el del día
    siguiente, que es el desfase que traen algunos archivos históricos).
    Un "0" en el reporte para uno de estos días no significa que no hubo
    actividad: significa que no había de dónde sacar el dato.
    """
    dias_con_archivo = set()
    for f in archivos:
        m = re.search(r"(\d{8})", os.path.basename(f))
        if not m:
            continue
        try:
            dias_con_archivo.add(datetime.strptime(m.group(1), "%Y%m%d").date())
        except ValueError:
            continue

    faltantes = []
    d = desde
    while d <= hasta:
        if d not in dias_con_archivo and (d + timedelta(days=1)) not in dias_con_archivo:
            faltantes.append(d)
        d += timedelta(days=1)
    return faltantes


def formatea_duracion(segundos):
    """Convierte segundos a un texto legible: '3s', '2m 05s', '1h 03m'."""
    segundos = max(0, int(round(segundos)))
    h, resto = divmod(segundos, 3600)
    m, s = divmod(resto, 60)
    if h:
        return f"{h}h {m:02d}m"
    if m:
        return f"{m}m {s:02d}s"
    return f"{s}s"


def imprime_barra_progreso(actual, total, tiempo_inicio, ancho=30):
    """Barra de progreso en la misma línea de la consola (sin dependencias),
    con tiempo transcurrido y tiempo restante estimado según el ritmo
    observado hasta el momento."""
    if total <= 0:
        return
    pct = actual / total
    llenado = int(ancho * pct)
    barra = "#" * llenado + "-" * (ancho - llenado)
    transcurrido = time.time() - tiempo_inicio
    if actual > 0:
        promedio_por_archivo = transcurrido / actual
        restante = promedio_por_archivo * (total - actual)
        eta_txt = formatea_duracion(restante) if actual < total else "0s"
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


def busca(usuario, desde, hasta, progreso=True):
    usuario = usuario.strip().lower()
    ini_dt = datetime.combine(desde, datetime.min.time(), tzinfo=TZ_CDMX)
    fin_dt = datetime.combine(hasta, datetime.max.time(), tzinfo=TZ_CDMX)

    archivos = archivos_candidatos(desde, hasta)
    if progreso:
        print(f"Revisando {len(archivos)} archivo(s) candidato(s)...")

    resultados = []
    vistos = set()
    duplicados = 0
    archivos_omitidos_estructura = []
    archivos_procesados_ok = 0
    archivos_validos = []
    total_archivos = len(archivos)
    tiempo_inicio_busqueda = time.time()
    for i, path in enumerate(archivos, 1):
        if progreso:
            imprime_barra_progreso(i, total_archivos, tiempo_inicio_busqueda)
        try:
            filas = lee_csv_robusto(path)
        except Exception as e:
            print(f"\n  [AVISO] No se pudo leer {os.path.basename(path)}: {e}")
            continue

        if not filas:
            continue

        # Validación de estructura: se espera EXACTAMENTE 8 columnas. Un
        # archivo con más o menos columnas (por ejemplo, exportaciones con
        # "ReceivedUTC"/"ToIP"/"TraceId" de más) se RECHAZA por completo en
        # lugar de leerlo a ciegas: si se procesara igual, fila[:8] tomaría
        # las columnas en la posición equivocada y mezclaría los datos sin
        # que se note (p. ej. una IP cayendo en el campo de tamaño).
        if len(filas[0]) != 8:
            archivos_omitidos_estructura.append(
                f"{os.path.basename(path)} ({len(filas[0])} columnas, se esperaban 8)"
            )
            continue
        archivos_procesados_ok += 1
        archivos_validos.append(path)

        for fila in filas:
            if len(fila) != 8:
                continue
            received, sender, recipient, subject, status, fromip, size, msgid = fila[
                :8
            ]
            if received.strip().lower() == "received":
                continue
            s_low, r_low = sender.strip().lower(), recipient.strip().lower()
            if usuario != s_low and usuario != r_low:
                continue
            dt_cdmx = parsea_fecha_cdmx(received)
            if dt_cdmx is None:
                continue
            if not (ini_dt <= dt_cdmx <= fin_dt):
                continue

            # Huella única del mensaje: si el mismo correo aparece más de
            # una vez entre los archivos (por ejemplo, dos exportaciones
            # que se traslapan en fechas), se cuenta una sola vez.
            huella = (msgid.strip(), received.strip(), s_low, r_low)
            if huella in vistos:
                duplicados += 1
                continue
            vistos.add(huella)

            resultados.append({
                "fecha_original": received.strip(),
                "fecha_cdmx": dt_cdmx,
                "fecha_utc": dt_cdmx.astimezone(TZ_UTC),
                "sender": sender,
                "recipient": recipient,
                "subject": subject,
                "status": status,
                "size": size,
                "msgid": msgid,
            })

    resultados.sort(key=lambda r: r["fecha_cdmx"])

    auditoria = {
        "archivos_candidatos": total_archivos,
        "archivos_procesados_ok": archivos_procesados_ok,
        "archivos_omitidos_estructura": archivos_omitidos_estructura,
        "dias_sin_cobertura": dias_sin_cobertura(desde, hasta, archivos_validos),
        "duplicados": duplicados,
        "tiempo_busqueda_seg": time.time() - tiempo_inicio_busqueda,
    }
    return resultados, auditoria


# LOGO BASE64 (INE, 432x118 px, PNG)
LOGO_INE_PNG_BASE64 = (
    "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"
    "Zm1D3l5+/twjcARAEAr5eMpx+r8HfD6NWtbSzNwQZUxk3n7rF1/fyL793AyNRVkZOf36/PDFZ7tFIgEmL23bO8OnWnA/q6JDOwr3"
    "/HffjXRpbC2TyXGOGDChQyuU42JSVyw/vmPbZUKY7PLE+SNGdETAOdn5P/9wtEe3pf5+kTnZzENbzGHRkPGpUIB8h4/oyOORqKiU"
    "vr1++HP7FYVC0bKVXdduLeRyBiKYUar5JOJHLzgKhQI9PT4Q8HnwYs6sHQF+URTFUKpUytgyNsX+hAtSYPRKJdh2w7qz2zZfhGxe"
    "fx6ESMG7UomWSK1P79YOKoLev9dz+fLjeKj1zVd///7bf+J86ZCh7cGJ4K8RQ1diOowt3cSETDs7c0KIlZUJIQSbhn26L4uJSR0/"
    "sTufTyOdXt2X7t55FbFj5jjknbZMOoieomBMqV5GRiKlUpmbkw+O3vXnVcSGU4WciV6BgHCiMqtXhzpLiMwoqm57DKu1o1nPYa4T"
    "5vXpM8q1adtGFjZGfCGPGXJm6Cn1C+ssvoCHLeqA+1HHNt0+s/v+k5thceGp0aEpPldCT2JxvflO4L0oqVgGdlNfVUxzJfOLLOhI"
    "ZCBo0MikWRvbvmM8xs7t0emt5tjihB7CmNXeG9Mc8Jq7u6OHhxNuA6GQ7+bm4O7u4OxsjaCcnK1Q5e7hKBDwNm2Z0f8tN1tb8wD/"
    "aBNj/X4D3H9bNQU2LVvbvTu0g0tjm7btnM3MDE3NDN09nNzcHRpYmaD2wxn93NwdsY3l6moPunl3WIe+/VyxBdauvcusOQObN2/k"
    "6uZgaWUMy1kfDWjt6gDL1q72IhF/0pSeM2b292jj7OBg4dHGqV9/t8VLx7i6OhRFTCQStGjQegAZAAAH6UlEQVTeyN3D0dm5AWDH"
    "2P2yYhJs0KOjo5WdvTmmtK1dHWwbmScnZPQf4NajZ0u4wgwO06LN2z/q1r154yY29vYWrVrbvzusAzbaQKbNmtkACjQHDc1fOHLy"
    "1N6ubvYO9hZYvA8a3HbTlpngFGyMeng4oiNTU0NDQz2wpJubQ9OmDR2drWZ+9Fbr1vbYWACTjh7TpbWrvZOjVV5uvoOjJaJyb+OE"
    "TIuKmakBbFzdHIAYAGSlXfvGiArZITV3d6cGDUyMjEULF4/u0rVZq1Z28bFpqB03ofuKXycbG4vmLxgxbXpfxIMuIO07NPn+p/F9"
    "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"
    "QVQDAAnXdiVxzQ1mAAAAAElFTkSuQmCC"
)


def _logo_imagen():
    try:
        return BytesIO(base64.b64decode(LOGO_INE_PNG_BASE64))
    except Exception:
        return None


def F(**k):
    return Font(name="Arial", **k)


# ============================================================
# CONSTRUCCIÓN DE HOJA DE DETALLE
# ============================================================
def crea_hoja_detalle(ws, titulo_pestana, lista_datos, usuario, desde, hasta):
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

    r = 1
    logo_stream = _logo_imagen()
    if logo_stream:
        img = XLImage(logo_stream)
        alto_orig, ancho_orig = img.height, img.width
        ancho_px = 240
        alto_px = int(ancho_px * alto_orig / ancho_orig) if ancho_orig else 65
        img.width = ancho_px
        img.height = alto_px
        ws.row_dimensions[1].height = 32
        ws.row_dimensions[2].height = 32
        ws.add_image(img, "B1")
        r = 3
    else:
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=LAST_COL)
        c = ws.cell(r, 2, "INSTITUTO NACIONAL ELECTORAL")
        c.font = F(bold=True, size=15, color=LILA_INST)
        r += 1

    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=LAST_COL)
    ws.cell(
        r, 2, "Unidad Técnica de Servicios de Informática (UTSI)"
    ).font = F(size=11)
    r += 1
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=LAST_COL)
    ws.cell(
        r,
        2,
        "Departamento de Soporte Técnico y Administración de Servicios de"
        " Colaboración (DSTyASC)",
    ).font = F(size=10, italic=True)
    r += 2

    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=LAST_COL)
    c = ws.cell(r, 2, f"DETALLE DE MENSAJES {titulo_pestana.upper()}")
    c.font = F(bold=True, size=14, color=BLANCO)
    c.fill = PatternFill("solid", fgColor=LILA_INST)
    c.alignment = Alignment(horizontal="center")
    ws.row_dimensions[r].height = 24
    r += 2

    headers = [
        "#",
        "FECHA ORIGINAL",
        "FECHA (CDMX / UTC-6)",
        "REMITENTE",
        "DESTINATARIO",
        "ASUNTO",
        "ESTATUS",
        "TAMAÑO (KB)",
    ]
    for j, h in enumerate(headers, start=1):
        c = ws.cell(r, j, h)
        c.font = F(bold=True, size=9, color=BLANCO)
        c.fill = PatternFill("solid", fgColor=LILA_OSCURO)
        c.alignment = Alignment(
            horizontal="center", vertical="center", wrap_text=True
        )
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
        vals = [
            idx,
            x["fecha_original"],
            x["fecha_cdmx"].strftime("%d/%m/%Y %H:%M:%S"),
            x["sender"],
            x["recipient"],
            x["subject"],
            st_es,
            kb,
        ]
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


# ============================================================
# REPORTE EXCEL MULTI-PESTAÑA
# ============================================================
def genera_excel(resultados, usuario, desde, hasta, out_path):
    thin = Side(style="thin", color="BFBFBF")
    border_all = Border(left=thin, right=thin, top=thin, bottom=thin)

    wb = openpyxl.Workbook()

    enviados_list = [
        x for x in resultados if x["sender"].strip().lower() == usuario.lower()
    ]
    recibidos_list = [
        x for x in resultados if x["recipient"].strip().lower() == usuario.lower()
    ]

    ws1 = wb.active
    ws1.title = "Reporte General"
    ws1.sheet_view.showGridLines = False

    LAST_COL = 7
    widths = [6, 20, 26, 26, 45, 14, 10]
    for i, w in enumerate(widths, start=1):
        ws1.column_dimensions[get_column_letter(i)].width = w

    r = 1
    logo_stream = _logo_imagen()
    if logo_stream:
        img = XLImage(logo_stream)
        alto_orig, ancho_orig = img.height, img.width
        ancho_px = 240
        alto_px = int(ancho_px * alto_orig / ancho_orig) if ancho_orig else 65
        img.width = ancho_px
        img.height = alto_px
        ws1.row_dimensions[1].height = 32
        ws1.row_dimensions[2].height = 32
        ws1.add_image(img, "B1")
        r = 3
    else:
        ws1.merge_cells(start_row=r, start_column=2, end_row=r, end_column=LAST_COL)
        c = ws1.cell(r, 2, "INSTITUTO NACIONAL ELECTORAL")
        c.font = F(bold=True, size=15, color=LILA_INST)
        r += 1

    ws1.merge_cells(start_row=r, start_column=2, end_row=r, end_column=LAST_COL)
    ws1.cell(
        r, 2, "Unidad Técnica de Servicios de Informática (UTSI)"
    ).font = F(size=11)
    r += 1
    ws1.merge_cells(start_row=r, start_column=2, end_row=r, end_column=LAST_COL)
    ws1.cell(
        r,
        2,
        "Departamento de Soporte Técnico y Administración de Servicios de"
        " Colaboración (DSTyASC)",
    ).font = F(size=10, italic=True)
    r += 2

    ws1.merge_cells(start_row=r, start_column=2, end_row=r, end_column=LAST_COL)
    c = ws1.cell(r, 2, "REPORTE DE BITÁCORA DE CORREO ELECTRÓNICO")
    c.font = F(bold=True, size=14, color=BLANCO)
    c.fill = PatternFill("solid", fgColor=LILA_INST)
    c.alignment = Alignment(horizontal="center")
    ws1.row_dimensions[r].height = 24
    r += 2

    def bloque_titulo(row, texto):
        ws1.merge_cells(
            start_row=row, start_column=2, end_row=row, end_column=LAST_COL
        )
        c = ws1.cell(row, 2, texto)
        c.font = F(bold=True, size=11, color=BLANCO)
        c.fill = PatternFill("solid", fgColor=LILA_OSCURO)
        c.alignment = Alignment(horizontal="left", indent=1, vertical="center")
        ws1.row_dimensions[row].height = 20

    bloque_titulo(r, "DATOS DE LA CONSULTA")
    r += 1
    datos = [
        ("Usuario consultado:", usuario),
        (
            "Periodo:",
            f"{desde.strftime('%d/%m/%Y')} al {hasta.strftime('%d/%m/%Y')}"
            " (horario CDMX)",
        ),
        (
            "Fecha de generación:",
            datetime.now(TZ_CDMX).strftime("%d/%m/%Y %H:%M hrs (CDMX)"),
        ),
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
    enviados = len(enviados_list)
    recibidos = len(recibidos_list)
    por_status = {}
    for x in resultados:
        st = status_label(x["status"])
        por_status[st] = por_status.get(st, 0) + 1

    bloque_titulo(r, "RESUMEN CONSOLIDADO")
    r += 1
    for etiqueta, valor in [
        ("Total de mensajes", total),
        ("Enviados por el usuario", enviados),
        ("Recibidos por el usuario", recibidos),
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

    # --- Desglose por periodo: por DÍA si el rango es corto, por MES si es
    # largo (arriba de un mes, listar 365 días sería inmanejable). ---
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
    _d = desde
    while _d <= hasta:
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
        MESES_ES = [
            "", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
            "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre",
        ]
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
        y, m = desde.year, desde.month
        while (y, m) <= (hasta.year, hasta.month):
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

    if por_status:
        ws1.cell(r, 2, "Desglose por estatus").font = F(
            bold=True, size=10, color=LILA_OSCURO
        )
        r += 1
        hdr = ws1.cell(r, 2, "Estatus")
        hdr.font = F(bold=True, size=9, color=BLANCO)
        hdr.fill = PatternFill("solid", fgColor=LILA_INST)
        ws1.merge_cells(start_row=r, start_column=2, end_row=r, end_column=5)
        hdr2 = ws1.cell(r, 6, "Cantidad")
        hdr2.font = F(bold=True, size=9, color=BLANCO)
        hdr2.fill = PatternFill("solid", fgColor=LILA_INST)
        ws1.merge_cells(start_row=r, start_column=6, end_row=r, end_column=LAST_COL)
        for col in range(2, LAST_COL + 1):
            ws1.cell(r, col).border = border_all
        r += 1
        for st, cnt in sorted(por_status.items(), key=lambda x: -x[1]):
            ws1.merge_cells(start_row=r, start_column=2, end_row=r, end_column=5)
            ws1.cell(r, 2, st).font = F(size=9)
            ws1.merge_cells(
                start_row=r, start_column=6, end_row=r, end_column=LAST_COL
            )
            c2 = ws1.cell(r, 6, cnt)
            c2.font = F(size=9)
            c2.alignment = Alignment(horizontal="center")
            low = st.lower()
            fill = (
                VERDE_OK
                if "entreg" in low
                else (ROJO_ERR if "error" in low else GRIS_CLARO)
            )
            for col in range(2, LAST_COL + 1):
                ws1.cell(r, col).fill = PatternFill("solid", fgColor=fill)
                ws1.cell(r, col).border = border_all
            r += 1
        r += 1

        raws_presentes = {}
        for x in resultados:
            raws_presentes[status_label(x["status"])] = x["status"]

        bloque_titulo(r, "NOMENCLATURA DE ESTATUS")
        r += 1
        for st_label in sorted(raws_presentes.keys()):
            desc = status_desc(raws_presentes[st_label])
            c1 = ws1.cell(r, 2, st_label)
            c1.font = F(bold=True, size=9, color=LILA_OSCURO)
            ws1.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
            c1.alignment = Alignment(vertical="top")
            c2 = ws1.cell(r, 4, desc)
            c2.font = F(size=9)
            ws1.merge_cells(
                start_row=r, start_column=4, end_row=r, end_column=LAST_COL
            )
            c2.alignment = Alignment(vertical="top", wrap_text=True)
            for col in range(2, LAST_COL + 1):
                ws1.cell(r, col).border = border_all
            ws1.row_dimensions[r].height = 26
            r += 1

    ws1.page_setup.orientation = "landscape"
    ws1.page_setup.fitToWidth = 1
    ws1.page_setup.fitToHeight = 0
    ws1.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws1.page_margins = PageMargins(left=0.4, right=0.4, top=0.5, bottom=0.5)
    ws1.print_area = f"A1:{get_column_letter(LAST_COL)}{r}"

    ws2 = wb.create_sheet(title="Enviados")
    crea_hoja_detalle(ws2, "Enviados", enviados_list, usuario, desde, hasta)

    ws3 = wb.create_sheet(title="Recibidos")
    crea_hoja_detalle(ws3, "Recibidos", recibidos_list, usuario, desde, hasta)

    wb.save(out_path)
    return total


def escribe_log_auditoria(usuario, desde, hasta, auditoria, total_encontrados, xlsx_path):
    """
    Escribe (agrega) el detalle técnico de la búsqueda a un log propio del
    usuario consultado, en OUT_DIR/Logs/<usuario>.log. El Excel/PDF entregable
    solo trae la información pedida; este log es el rastro auditable de
    cómo se hizo la búsqueda: qué se revisó, qué se omitió y por qué.
    """
    log_dir = os.path.join(OUT_DIR, "Logs")
    os.makedirs(log_dir, exist_ok=True)

    safe_user = re.sub(r"[^a-zA-Z0-9]", "_", usuario.strip().lower())
    log_path = os.path.join(log_dir, f"{safe_user}.log")

    omitidos = auditoria["archivos_omitidos_estructura"]
    huecos = auditoria["dias_sin_cobertura"]

    lineas = []
    lineas.append("=" * 78)
    lineas.append(f"Auditoría de búsqueda — {datetime.now(TZ_CDMX).strftime('%d/%m/%Y %H:%M:%S hrs (CDMX)')}")
    lineas.append(f"Usuario consultado : {usuario}")
    lineas.append(
        f"Periodo consultado : {desde.strftime('%d/%m/%Y')} al {hasta.strftime('%d/%m/%Y')} (horario CDMX)"
    )
    lineas.append("-" * 78)
    lineas.append(
        f"Archivos fuente revisados        : {auditoria['archivos_procesados_ok']} de"
        f" {auditoria['archivos_candidatos']} candidatos"
    )
    lineas.append(
        f"Archivos con estructura no estándar (omitidos) : {len(omitidos) if omitidos else 'Ninguno'}"
    )
    for nombre in omitidos:
        lineas.append(f"    - {nombre}")
    lineas.append(
        f"Días del rango sin archivo fuente : {len(huecos) if huecos else 'Ninguno — cobertura completa'}"
    )
    if huecos:
        lineas.append("    " + ", ".join(d.strftime("%d/%m/%Y") for d in huecos))
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
        # Selecciona TODAS las hojas para que el PDF incluya las 3 pestañas
        # (si no se seleccionan todas, ExportAsFixedFormat solo exporta la
        # hoja activa).
        wb.Sheets.Select()
        wb.ActiveSheet.ExportAsFixedFormat(0, pdf_path)
        wb.Close(False)
        excel.Quit()
        return True
    except Exception as e:
        print(f"[AVISO] No se pudo generar el PDF automáticamente: {e}")
        print(
            "        Abre el .xlsx en Excel y usa 'Guardar como > PDF' manualmente."
        )
        return False


# ============================================================
# MAIN
# ============================================================
def main():
    global CSV_DIR
    tiempo_inicio_total = time.time()
    ap = argparse.ArgumentParser(
        description="Busca bitácoras de correo por usuario y fecha."
    )
    ap.add_argument("--usuario", help="Correo del usuario a buscar")
    ap.add_argument(
        "--desde", help="Fecha inicio, formato yyyy-mm-dd (horario CDMX)"
    )
    ap.add_argument(
        "--hasta", help="Fecha fin, formato yyyy-mm-dd (horario CDMX)"
    )
    ap.add_argument(
        "--csv-dir",
        dest="csv_dir",
        default=None,
        help=(
            "Carpeta donde están los MessageTrace_*.csv. Por defecto lee "
            f"del servidor origen por red: {CSV_DIR}. Usa este parámetro "
            r"para leer de otro lado, por ejemplo E:\Bitacoras (la copia"
            " local), si el servidor remoto no está disponible."
        ),
    )
    args = ap.parse_args()

    if args.csv_dir:
        CSV_DIR = args.csv_dir
        print(f"Leyendo bitácoras desde: {CSV_DIR}")

    usuario = args.usuario or input("Usuario (correo): ").strip()
    desde_s = args.desde or input("Fecha inicio (yyyy-mm-dd): ").strip()
    hasta_s = args.hasta or input("Fecha fin    (yyyy-mm-dd): ").strip()

    try:
        desde = datetime.strptime(desde_s, "%Y-%m-%d").date()
        hasta = datetime.strptime(hasta_s, "%Y-%m-%d").date()
    except ValueError:
        print(
            "ERROR: las fechas deben tener formato yyyy-mm-dd, por ejemplo"
            " 2026-06-01"
        )
        sys.exit(1)

    if hasta < desde:
        print("ERROR: la fecha 'hasta' es anterior a la fecha 'desde'.")
        sys.exit(1)

    print(
        f"\nBuscando bitácoras de {usuario} del {desde} al {hasta} (horario"
        " CDMX)..."
    )
    resultados, auditoria = busca(usuario, desde, hasta)
    print(f"Encontrados: {len(resultados)} mensaje(s) único(s).")
    print(
        f"Archivos revisados: {auditoria['archivos_procesados_ok']} de"
        f" {auditoria['archivos_candidatos']} candidatos."
    )
    if auditoria["archivos_omitidos_estructura"]:
        print(
            f"[ATENCIÓN] {len(auditoria['archivos_omitidos_estructura'])} archivo(s)"
            " con estructura distinta a la esperada (NO se procesaron, revísalos):"
        )
        for nombre in auditoria["archivos_omitidos_estructura"]:
            print(f"    - {nombre}")
    if auditoria["dias_sin_cobertura"]:
        dias_txt = ", ".join(d.strftime("%d/%m/%Y") for d in auditoria["dias_sin_cobertura"])
        print(
            f"[ATENCIÓN] {len(auditoria['dias_sin_cobertura'])} día(s) del rango sin"
            f" ningún archivo fuente disponible: {dias_txt}"
        )
    else:
        print("Cobertura de archivos: completa (todos los días del rango tienen fuente).")
    if auditoria["duplicados"]:
        print(f"Se detectaron y omitieron {auditoria['duplicados']} registro(s) duplicado(s) entre los archivos.")
    print(f"Tiempo de búsqueda: {formatea_duracion(auditoria['tiempo_busqueda_seg'])}")

    if not resultados:
        print("No se generará reporte: no hubo resultados para ese usuario/rango.")
        log_path = escribe_log_auditoria(usuario, desde, hasta, auditoria, 0, "N/A (sin resultados, no se generó reporte)")
        print(f"Auditoría de la búsqueda agregada al log: {log_path}")
        print(f"Tiempo total de ejecución: {formatea_duracion(time.time() - tiempo_inicio_total)}")
        sys.exit(0)

    safe_user = re.sub(r"[^a-zA-Z0-9]", "_", usuario)
    nombre_base = f"Bitacora_{safe_user}_{desde}_a_{hasta}"
    xlsx_path = os.path.join(OUT_DIR, nombre_base + ".xlsx")
    pdf_path = os.path.join(OUT_DIR, nombre_base + ".pdf")

    tiempo_inicio_excel = time.time()
    genera_excel(resultados, usuario, desde, hasta, xlsx_path)
    print(f"Excel generado con 3 pestañas ({formatea_duracion(time.time() - tiempo_inicio_excel)}): {xlsx_path}")

    log_path = escribe_log_auditoria(usuario, desde, hasta, auditoria, len(resultados), xlsx_path)
    print(f"Auditoría de la búsqueda agregada al log: {log_path}")

    tiempo_inicio_pdf = time.time()
    if exporta_pdf(xlsx_path, pdf_path):
        print(f"PDF generado ({formatea_duracion(time.time() - tiempo_inicio_pdf)}): {pdf_path}")

    print(f"Tiempo total de ejecución: {formatea_duracion(time.time() - tiempo_inicio_total)}")


if __name__ == "__main__":
    main()
