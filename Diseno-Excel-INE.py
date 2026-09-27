#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Diseno-Excel-INE.py
=====================
Solo el DISEÑO de Excel (paleta, encabezado institucional, barra de
título, tabla con cebra/zebra y bordes, bloque de resumen con colores por
estatus) separado de la lógica de búsqueda de correos, para poder
reutilizarlo en cualquier otro proyecto/reporte.

Cómo usarlo en otro script:
    1. Copia este archivo a tu proyecto (o solo las funciones que
       necesites).
    2. Arma tu propio openpyxl.Workbook() y usa las funciones de abajo
       para pintar cada parte: encabezado_institucional(), titulo_seccion(),
       tabla_con_zebra(), oculta_no_usado().
    3. Cambia los colores en la sección PALETA si es para otra
       institución/marca.

Corre este archivo directo (`python Diseno-Excel-INE.py`) y genera
`Demo-Diseno.xlsx` con un ejemplo de cada pieza, para que abras el Excel y
veas exactamente cómo se ve cada función que uses. Pásale la ruta de un
logo como argumento para probar esa pieza también:
    python Diseno-Excel-INE.py ruta/al/logo.png

Requisitos: pip install openpyxl pillow
(pillow solo es necesaria si vas a usar logo_path/inserta_logo(); sin ella
el encabezado simplemente cae en el texto).
"""

import os

from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins
from openpyxl.worksheet.properties import PageSetupProperties

# ============================================================
# PALETA — cambia estos 6 valores para adaptarlo a otra marca
# ============================================================
COLOR_PRINCIPAL = "674092"   # morado INE: títulos, barra de estatus
COLOR_OSCURO = "49276F"      # encabezados de tabla, subtítulos
COLOR_GRIS_CLARO = "F2F2F2"  # renglones "cebra" (alternados)
COLOR_BLANCO = "FFFFFF"      # texto sobre fondo oscuro
COLOR_VERDE_OK = "E2EFDA"    # fila de estatus "positivo" (ej. Entregado)
COLOR_ROJO_ERR = "FCE4E4"    # fila de estatus "de error"

BORDE_FINO = Side(style="thin", color="BFBFBF")
BORDE_CELDA = Border(left=BORDE_FINO, right=BORDE_FINO, top=BORDE_FINO, bottom=BORDE_FINO)
BORDE_INFERIOR = Border(bottom=BORDE_FINO)


def fuente(**kwargs):
    """Fuente institucional: Arial en todo el libro. Ejemplo:
    fuente(bold=True, size=14, color=COLOR_BLANCO)"""
    return Font(name="Arial", **kwargs)


# ============================================================
# PIEZA 1a: Logo (opcional). Se valida con PIL ANTES de insertarlo en el
# Excel: una imagen corrupta o con datos truncados suele "decodificar" el
# encabezado sin error pero se ve como un cuadro negro/gris sólido al
# abrir el archivo. Forzar la carga completa aquí evita ese problema —
# si falla, no se inserta nada (deja hueco para el texto en su lugar).
# ============================================================
def inserta_logo(ws, image_path, cell="B1", ancho_px=240, filas_alto=(1, 2), alto_fila_px=32):
    """
    Inserta una imagen (logo) anclada en `cell`, escalada a `ancho_px` de
    ancho manteniendo la proporción. Regresa True si se insertó, False si
    no (archivo inexistente, formato no soportado, imagen corrupta, etc.)
    — en ese caso el llamador debe usar el texto de respaldo en su lugar.
    """
    if not image_path or not os.path.isfile(image_path):
        return False
    try:
        from PIL import Image as PILImage
        with PILImage.open(image_path) as pil_img:
            pil_img.load()  # fuerza la decodificación completa, no solo el header
            ancho_orig, alto_orig = pil_img.size
    except Exception:
        return False

    try:
        img = XLImage(image_path)
        alto_px = int(ancho_px * alto_orig / ancho_orig) if ancho_orig else 65
        img.width = ancho_px
        img.height = alto_px
        for fila in filas_alto:
            ws.row_dimensions[fila].height = alto_fila_px
        ws.add_image(img, cell)
        return True
    except Exception:
        return False


# ============================================================
# PIEZA 1b: Encabezado institucional (logo opcional + 2 líneas de texto,
# o 3 líneas de texto si no hay logo o no se pudo insertar).
# ============================================================
def encabezado_institucional(ws, last_col, logo_path=None,
                              linea1="INSTITUTO NACIONAL ELECTORAL",
                              linea2="Unidad Técnica de Servicios de Informática (UTSI)",
                              linea3="Departamento de Soporte Técnico y Administración de Servicios de"
                                     " Colaboración (DSTyASC)"):
    """
    Dibuja el encabezado a partir de la fila 1, columna B (deja la
    columna A libre como margen visual). Si `logo_path` apunta a una
    imagen válida, la inserta y omite `linea1` (el logo ya trae el
    nombre); si no hay logo o falla la validación, cae en las 3 líneas de
    texto. Regresa el número de fila donde puede continuar el resto del
    contenido.
    """
    r = 1
    if logo_path and inserta_logo(ws, logo_path, cell=f"B{r}"):
        r = 3
    else:
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=last_col)
        c = ws.cell(r, 2, linea1)
        c.font = fuente(bold=True, size=15, color=COLOR_PRINCIPAL)
        r += 1

    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=last_col)
    ws.cell(r, 2, linea2).font = fuente(size=11)
    r += 1

    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=last_col)
    ws.cell(r, 2, linea3).font = fuente(size=10, italic=True)
    r += 2  # deja una fila en blanco después del encabezado
    return r


# ============================================================
# PIEZA 2: Barra de título de sección (fondo morado, texto blanco)
# ============================================================
def titulo_seccion(ws, row, last_col, texto, tamano=14, fondo=None, centrado=True, alto_fila=24):
    """
    Barra de título coloreada de ancho completo. Úsala tanto para el
    título grande de la hoja como para subtítulos de bloque (con
    tamano=11 y fondo=COLOR_OSCURO, por ejemplo).
    """
    fondo = fondo or COLOR_PRINCIPAL
    ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=last_col)
    c = ws.cell(row, 2, texto)
    c.font = fuente(bold=True, size=tamano, color=COLOR_BLANCO)
    c.fill = PatternFill("solid", fgColor=fondo)
    c.alignment = Alignment(horizontal="center" if centrado else "left", indent=0 if centrado else 1, vertical="center")
    ws.row_dimensions[row].height = alto_fila
    return row + 1


# ============================================================
# PIEZA 3: Bloque "etiqueta: valor" (para secciones de datos/resumen)
# ============================================================
def fila_etiqueta_valor(ws, row, last_col, etiqueta, valor, col_split=3,
                         valor_destacado=False):
    """
    Escribe una fila 'Etiqueta:  Valor'. `col_split` es la última columna
    que ocupa la etiqueta (el valor empieza en col_split+1). Con
    valor_destacado=True el valor se ve grande y en el color principal
    (para un número importante, como un total).
    """
    c1 = ws.cell(row, 2, etiqueta)
    c1.font = fuente(bold=True, size=10)
    ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=col_split)
    ws.merge_cells(start_row=row, start_column=col_split + 1, end_row=row, end_column=last_col)
    c2 = ws.cell(row, col_split + 1, valor)
    if valor_destacado:
        c2.font = fuente(size=13, bold=True, color=COLOR_PRINCIPAL)
        c2.alignment = Alignment(horizontal="center")
        for col in range(2, last_col + 1):
            ws.cell(row, col).border = BORDE_CELDA
        ws.row_dimensions[row].height = 20
    else:
        c2.font = fuente(size=10)
    return row + 1


# ============================================================
# PIEZA 4: Tabla con encabezado morado, bordes y filas "cebra"
# ============================================================
def tabla_con_zebra(ws, start_row, headers, filas, anchos=None,
                     columnas_centradas=None, color_fila_fn=None):
    """
    Dibuja una tabla: fila de encabezado (fondo oscuro, texto blanco) y
    debajo cada fila de datos alternando fondo gris claro (cebra).

    - headers: lista de textos de columna, p. ej. ["#", "Nombre", "Monto"]
    - filas: lista de listas/tuplas con los valores de cada fila, en el
      mismo orden que headers
    - anchos: lista opcional de anchos de columna (en unidades de Excel)
    - columnas_centradas: set opcional de índices (1-based) que se
      centran en vez de alinearse a la izquierda
    - color_fila_fn: función opcional fila -> color_hex (6 caracteres) o
      None, para colorear una fila completa (p. ej. verde si "status"=="OK",
      rojo si "status"=="Error"). Se aplica ADEMÁS del cebra.

    Regresa (fila_encabezado, ultima_fila) por si necesitas fijar paneles
    o el área de impresión con esos números.
    """
    last_col = len(headers)
    if anchos:
        for i, w in enumerate(anchos, start=1):
            ws.column_dimensions[get_column_letter(i)].width = w
    columnas_centradas = columnas_centradas or set()

    r = start_row
    for j, h in enumerate(headers, start=1):
        c = ws.cell(r, j, h)
        c.font = fuente(bold=True, size=9, color=COLOR_BLANCO)
        c.fill = PatternFill("solid", fgColor=COLOR_OSCURO)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDE_CELDA
    ws.row_dimensions[r].height = 22
    fila_encabezado = r
    r += 1

    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")
    fill_zebra = PatternFill("solid", fgColor=COLOR_GRIS_CLARO)

    for idx, fila in enumerate(filas, start=1):
        zebra = idx % 2 == 0
        color_extra = color_fila_fn(fila) if color_fila_fn else None
        for j, valor in enumerate(fila, start=1):
            c = ws.cell(r, j, valor)
            c.font = fuente(size=8)
            c.alignment = align_center if j in columnas_centradas else align_left
            c.border = BORDE_INFERIOR
            if color_extra:
                c.fill = PatternFill("solid", fgColor=color_extra)
            elif zebra:
                c.fill = fill_zebra
        r += 1

    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_margins = PageMargins(left=0.4, right=0.4, top=0.5, bottom=0.5)
    ws.print_area = f"A1:{get_column_letter(last_col)}{r - 1}"
    ws.print_title_rows = f"{fila_encabezado - 1}:{fila_encabezado - 1}"
    ws.freeze_panes = ws.cell(fila_encabezado, 1)
    return fila_encabezado, r - 1


# ============================================================
# PIEZA 5: Oculta columnas/filas sobrantes (para que la hoja se vea
# "acotada" a los datos reales al abrirse, sin espacio vacío de más)
# ============================================================
def oculta_no_usado(ws, last_col, last_row, col_buffer=40, row_buffer=300):
    for c in range(last_col + 1, last_col + 1 + col_buffer):
        ws.column_dimensions[get_column_letter(c)].hidden = True
    for rr in range(last_row + 1, last_row + 1 + row_buffer):
        ws.row_dimensions[rr].hidden = True


# ============================================================
# DEMO: genera un Excel de ejemplo usando las 5 piezas de arriba
# ============================================================
def _demo():
    import sys
    logo_path = sys.argv[1] if len(sys.argv) > 1 else None

    wb = Workbook()
    ws = wb.active
    ws.title = "Demo"
    ws.sheet_view.showGridLines = False

    LAST_COL = 7
    r = encabezado_institucional(ws, LAST_COL, logo_path=logo_path)
    r = titulo_seccion(ws, r, LAST_COL, "REPORTE DE EJEMPLO")
    r += 1

    r = titulo_seccion(ws, r, LAST_COL, "DATOS DE LA CONSULTA", tamano=11,
                        fondo=COLOR_OSCURO, centrado=False, alto_fila=20)
    r = fila_etiqueta_valor(ws, r, LAST_COL, "Generado por:", "Diseno-Excel-INE.py (demo)")
    r = fila_etiqueta_valor(ws, r, LAST_COL, "Total de registros:", 3, valor_destacado=True)
    r += 1

    headers = ["#", "Nombre", "Categoría", "Estatus", "Monto"]
    filas = [
        [1, "Elemento A", "Categoría 1", "OK", 1500],
        [2, "Elemento B", "Categoría 2", "Error", 320],
        [3, "Elemento C", "Categoría 1", "OK", 980],
    ]

    def color_por_estatus(fila):
        estatus = fila[3]
        if estatus == "OK":
            return COLOR_VERDE_OK
        if estatus == "Error":
            return COLOR_ROJO_ERR
        return None

    ws.column_dimensions["A"].width = 2
    fila_enc, ultima_fila = tabla_con_zebra(
        ws, r, headers, filas,
        anchos=[4, 22, 18, 12, 12],
        columnas_centradas={1, 4, 5},
        color_fila_fn=color_por_estatus,
    )
    oculta_no_usado(ws, LAST_COL, ultima_fila)

    wb.save("Demo-Diseno.xlsx")
    print("Generado: Demo-Diseno.xlsx — ábrelo para ver el resultado de cada pieza.")


if __name__ == "__main__":
    _demo()
