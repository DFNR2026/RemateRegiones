"""
Módulo 5: Reporte Final

Expone dos funciones públicas:
  actualizar_historial(causas) → APPEND a hoja CAUSAS de causas_ojv.xlsx
  generar_reporte(causas)      → crea Reporte_YYYY-MM-DD.xlsx

Input:  lista de causas con todos los campos de M1-M4
Output 1: causas_ojv.xlsx (hoja CAUSAS actualizada — historial deduplicación)
Output 2: Reporte_YYYY-MM-DD.xlsx con 3 pestañas y formato condicional
"""

import os
import datetime
import logging

import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils import get_column_letter

from config import CAUSAS_XLSX, BASE_DIR, REPORTES_DIR, SHEET_CAUSAS

logging.basicConfig(level=logging.INFO, format="%(asctime)s [M5] %(message)s")
log = logging.getLogger(__name__)


# ─────────────────────────────────────────────────────────────────
# Clasificación (sin tasación — Módulo 4 abandonado)
# ─────────────────────────────────────────────────────────────────

def _clasificar(causa: dict) -> str:
    """
    Clasifica cada causa según el estado de extracción de deuda:
      - CON DEUDA EXTRAÍDA : M3 encontró el monto en el PDF
      - SIN PDF            : M2 no descargó documento (OJV fallo, proc. no aplicable, etc.)
      - SIN MONTO EN PDF   : M2 descargó PDF pero M3 no pudo extraer el monto
    """
    if causa.get("monto_deuda_clp"):
        return "CON DEUDA EXTRAÍDA"
    if not causa.get("descargado"):
        return "SIN PDF"
    return "SIN MONTO EN PDF"


def _enriquecer_clasificacion(causas: list[dict]) -> list[dict]:
    """Agrega _clasificacion a cada causa (campo temporal de trabajo)."""
    for c in causas:
        c["_clasificacion"] = _clasificar(c)
    return causas


# Orden geográfico norte a sur de las Cortes de Apelaciones de regiones.
# Las cortes de RM (Santiago y San Miguel) van en su propia pestaña y no
# aparecen aquí.
_ORDEN_CORTES_NORTE_SUR = [
    "C.A. de Arica",
    "C.A. de Iquique",
    "C.A. de Antofagasta",
    "C.A. de Copiapó",
    "C.A. de La Serena",
    "C.A. de Valparaíso",
    "C.A. de Rancagua",
    "C.A. de Talca",
    "C.A. de Chillán",
    "C.A. de Concepción",
    "C.A. de Temuco",
    "C.A. de Valdivia",
    "C.A. de Puerto Montt",
    "C.A. de Coyhaique",
    "C.A. de Punta Arenas",
]
_IDX_CORTE = {corte: i for i, corte in enumerate(_ORDEN_CORTES_NORTE_SUR)}


def _ordenar_regiones(causas: list[dict]) -> list[dict]:
    """
    Ordena causas por monto_deuda_clp ascendente (menor deuda primero).
    Causas con deuda nula o 0 van al final, ordenadas por corte geográfica.
    """
    n = len(_ORDEN_CORTES_NORTE_SUR)

    def key(c):
        deuda = c.get("monto_deuda_clp") or 0
        if deuda > 0:
            return (0, deuda, 0)
        # Sin deuda: al fondo, ordenadas geográficamente norte→sur
        idx_corte = _IDX_CORTE.get(c.get("corte", ""), n)
        return (1, 0, idx_corte)

    return sorted(causas, key=key)


# ─────────────────────────────────────────────────────────────────
# Estilos reutilizables
# ─────────────────────────────────────────────────────────────────

_FILL_HEADER  = PatternFill("solid", fgColor="1F4E79")   # azul oscuro
_FILL_ODD     = PatternFill("solid", fgColor="F2F7FF")   # azul muy claro
_FILL_EVEN    = PatternFill("solid", fgColor="FFFFFF")   # blanco

_FONT_HEADER  = Font(bold=True, color="FFFFFF", size=11)
_FONT_BODY    = Font(size=10)
_FONT_BOLD    = Font(bold=True, size=10)

_ALIGN_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=False)
_ALIGN_LEFT   = Alignment(horizontal="left",   vertical="center", wrap_text=False)
_ALIGN_RIGHT  = Alignment(horizontal="right",  vertical="center", wrap_text=False)

_BORDER_THIN  = Border(
    bottom=Side(style="thin", color="CCCCCC"),
    right= Side(style="thin", color="CCCCCC"),
)

# Colores de clasificación (fondo)
_FILL_CON_DEUDA    = PatternFill("solid", fgColor="00B050")   # verde
_FILL_SIN_PDF      = PatternFill("solid", fgColor="FFC000")   # amarillo/ámbar
_FILL_SIN_MONTO    = PatternFill("solid", fgColor="FF0000")   # rojo

# Fuente para clasificaciones
_FONT_WHITE = Font(bold=True, color="FFFFFF", size=10)
_FONT_DARK  = Font(bold=True, color="333333", size=10)


# ─────────────────────────────────────────────────────────────────
# Estructura de columnas
# ─────────────────────────────────────────────────────────────────

# (header_label, campo_dict, ancho_col, formato_num)
_COLUMNAS = [
    ("ROL",             "rol",                 9,  None),
    ("Año",             "año",                 6,  None),
    ("Corte",           "corte",              20,  None),
    ("Tribunal",        "tribunal",           30,  None),
    ("Demandante",      "demandante",         26,  None),
    ("Demandado",       "demandado",          26,  None),
    ("Dirección",       "direccion",          36,  None),
    ("Comuna",          "comuna",             16,  None),
    ("Tipo Proc.",      "tipo_procedimiento", 13,  None),
    ("Deuda (CLP)",     "monto_deuda_clp",    17,  '#,##0'),
    ("Fechas Public.",  "fechas_publicacion", 18,  None),
    ("Motivo Fallo",    "motivo_fallo",       32,  None),
]


# ─────────────────────────────────────────────────────────────────
# Escritura de hoja de datos
# ─────────────────────────────────────────────────────────────────

def _escribir_hoja_datos(wb: openpyxl.Workbook, causas: list[dict], nombre: str) -> None:
    """
    Crea una hoja con la lista de causas ordenadas, headers y formato condicional.
    """
    ws = wb.create_sheet(nombre)

    # ── Fila de encabezados ──
    for col_idx, (header, _, ancho, _fmt) in enumerate(_COLUMNAS, 1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.fill      = _FILL_HEADER
        cell.font      = _FONT_HEADER
        cell.alignment = _ALIGN_CENTER
        cell.border    = _BORDER_THIN
        ws.column_dimensions[get_column_letter(col_idx)].width = ancho

    ws.row_dimensions[1].height = 22
    ws.freeze_panes = "A2"  # congelar fila de encabezados

    # ── Filas de datos ──
    for row_idx, causa in enumerate(causas, 2):
        fill = _FILL_ODD if row_idx % 2 == 1 else _FILL_EVEN

        for col_idx, (_, campo, _, fmt_num) in enumerate(_COLUMNAS, 1):
            valor = causa.get(campo, "")

            if valor is None:
                valor = ""
            # Monto monetario: 0 o None → celda en blanco
            elif campo == "monto_deuda_clp" and not valor:
                valor = ""
            # Corte: eliminar prefijo "C.A. de " para ahorrar espacio
            elif campo == "corte" and isinstance(valor, str) and valor.startswith("C.A. de "):
                valor = valor[8:]

            cell = ws.cell(row=row_idx, column=col_idx, value=valor)
            cell.fill      = fill
            cell.font      = _FONT_BODY
            cell.alignment = _ALIGN_RIGHT if fmt_num else _ALIGN_LEFT
            cell.border    = _BORDER_THIN

            if fmt_num and valor != "":
                cell.number_format = fmt_num

        ws.row_dimensions[row_idx].height = 16

    if not causas:
        return

    max_row = len(causas) + 1

    # ── Auto-filtro ──
    ws.auto_filter.ref = f"A1:{get_column_letter(len(_COLUMNAS))}{max_row}"


# ─────────────────────────────────────────────────────────────────
# Hoja Resumen
# ─────────────────────────────────────────────────────────────────

def _escribir_hoja_resumen(wb: openpyxl.Workbook, causas: list[dict]) -> None:
    """Crea la pestaña Resumen con estadísticas de la ejecución."""
    ws = wb.create_sheet("Resumen")
    ws.sheet_properties.tabColor = "1F4E79"

    # Anchos de columna
    ws.column_dimensions["A"].width = 32
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 30

    fecha_hoy = datetime.date.today().strftime("%d/%m/%Y")

    # Contadores
    total = len(causas)

    por_clas = {k: 0 for k in ("CON DEUDA EXTRAÍDA", "SIN PDF", "SIN MONTO EN PDF")}
    for c in causas:
        k = c.get("_clasificacion", "SIN PDF")
        por_clas[k] = por_clas.get(k, 0) + 1

    desc   = sum(1 for c in causas if c.get("descargado"))
    mand   = sum(1 for c in causas if c.get("tipo_documento") == "mandamiento")
    bases  = sum(1 for c in causas if c.get("tipo_documento") == "bases_remate")
    sin_d  = sum(1 for c in causas if not c.get("descargado"))

    # ── Helpers de escritura ──
    def titulo(row, texto):
        cell = ws.cell(row=row, column=1, value=texto)
        cell.font = Font(bold=True, color="FFFFFF", size=12)
        cell.fill = _FILL_HEADER
        cell.alignment = _ALIGN_LEFT
        ws.merge_cells(f"A{row}:C{row}")
        ws.row_dimensions[row].height = 20

    def fila(row, label, valor, nota="", fill=None):
        c1 = ws.cell(row=row, column=1, value=label)
        c2 = ws.cell(row=row, column=2, value=valor)
        c3 = ws.cell(row=row, column=3, value=nota)
        c1.font = c3.font = _FONT_BODY
        c2.font = _FONT_BOLD
        c2.alignment = _ALIGN_RIGHT
        if fill:
            c1.fill = c2.fill = c3.fill = fill
        ws.row_dimensions[row].height = 16

    def espacio(row):
        ws.row_dimensions[row].height = 8

    # ── Título ──
    r = 1
    ws.merge_cells(f"A{r}:C{r}")
    cell = ws.cell(row=r, column=1, value="RESUMEN — ANÁLISIS DE REMATES JUDICIALES")
    cell.font = Font(bold=True, color="FFFFFF", size=14)
    cell.fill = PatternFill("solid", fgColor="0D2D4F")
    cell.alignment = _ALIGN_CENTER
    ws.row_dimensions[r].height = 28

    r += 1
    fila(r, "Fecha de procesamiento:", fecha_hoy)

    r += 1; espacio(r)
    r += 1; titulo(r, "TOTALES")
    r += 1; fila(r, "Total causas procesadas",   total)

    r += 1; espacio(r)
    r += 1; titulo(r, "DESGLOSE POR ESTADO DE DEUDA")

    _filas_clas = [
        ("CON DEUDA EXTRAÍDA", "PDF descargado y monto encontrado",      _FILL_CON_DEUDA, _FONT_WHITE),
        ("SIN PDF",            "no descargado / OJV fallo / proc. N/A",  _FILL_SIN_PDF,   _FONT_DARK),
        ("SIN MONTO EN PDF",   "PDF descargado pero monto no extraído",  _FILL_SIN_MONTO, _FONT_WHITE),
    ]
    for clas, nota, fill, font in _filas_clas:
        r += 1
        fila(r, f"  {clas}", por_clas.get(clas, 0), nota, fill=fill)
        for col in (1, 2, 3):
            ws.cell(row=r, column=col).font = font

    r += 1; espacio(r)
    r += 1; titulo(r, "DOCUMENTOS DESCARGADOS (OJV)")
    r += 1; fila(r, "  Mandamientos (ejecutivo)",          mand)
    r += 1; fila(r, "  Bases de Remate (ley de bancos)",   bases)
    r += 1; fila(r, "  No descargados",                    sin_d)
    r += 1; fila(r, "  Total descargados",                 desc)



# ─────────────────────────────────────────────────────────────────
# FUNCIÓN PÚBLICA 1: actualizar_historial
# ─────────────────────────────────────────────────────────────────

def actualizar_historial(causas: list[dict]) -> None:
    """
    Agrega las causas procesadas esta semana al historial acumulativo
    en la hoja CAUSAS de causas_ojv.xlsx.

    Solo hace APPEND — nunca borra filas existentes.
    Evita duplicar ROLes que ya estén en el historial.

    Args:
        causas: lista de dicts (output del pipeline completo)
    """
    log.info(f"Actualizando historial — {len(causas)} causa(s)")

    wb = openpyxl.load_workbook(CAUSAS_XLSX)

    if SHEET_CAUSAS not in wb.sheetnames:
        ws = wb.create_sheet(SHEET_CAUSAS)
        ws.append(["ROL", "AÑO", "CORTE", "TRIBUNAL", "FECHA_PROCESADO"])
        log.warning(f"Hoja '{SHEET_CAUSAS}' no existía — se creó")
    else:
        ws = wb[SHEET_CAUSAS]
        # Agregar cabecera FECHA_PROCESADO si el sheet existente no la tiene
        if ws.cell(1, 5).value != "FECHA_PROCESADO":
            ws.cell(1, 5).value = "FECHA_PROCESADO"

    # Leer ROLes ya existentes para evitar duplicados
    roles_existentes: set[tuple] = set()
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0] and row[1]:
            roles_existentes.add((str(row[0]).strip(), str(row[1]).strip()))

    fecha_hoy = datetime.date.today().isoformat()
    agregados = 0

    for c in causas:
        clave = (str(c.get("rol", "")).strip(), str(c.get("año", "")).strip())
        if not clave[0]:
            continue
        if clave in roles_existentes:
            continue

        ws.append([
            c.get("rol", ""),
            c.get("año", ""),
            c.get("corte", ""),
            c.get("tribunal", ""),
            fecha_hoy,
        ])
        roles_existentes.add(clave)
        agregados += 1

    wb.save(CAUSAS_XLSX)
    log.info(f"Historial actualizado: {agregados} causas nuevas agregadas")


# ─────────────────────────────────────────────────────────────────
# FUNCIÓN PÚBLICA 2: generar_reporte
# ─────────────────────────────────────────────────────────────────

def generar_reporte(causas: list[dict]) -> str:
    """
    Genera el reporte semanal Reporte_YYYY-MM-DD.xlsx con:
      - Pestaña "Regiones" (todas las causas, ordenadas geográficamente por ratio)
      - Pestaña "Resumen"  (estadísticas y notas)

    Args:
        causas: lista de dicts (output del pipeline completo)

    Returns:
        Ruta del archivo generado.
    """
    log.info(f"Generando reporte — {len(causas)} causa(s)")

    # Usar nombre completo del demandado (OJV Litigantes) si está disponible
    for c in causas:
        if c.get("demandado_nombre"):
            c["demandado"] = c["demandado_nombre"]

    # Clasificar por estado de deuda
    _enriquecer_clasificacion(causas)

    # Ordenar por corte geográfica norte→sur y luego por ratio
    regiones = _ordenar_regiones(causas)

    log.info(f"  Regiones: {len(regiones)}")

    # Crear workbook
    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # quitar hoja vacía por defecto

    _escribir_hoja_datos(wb, regiones, "Regiones")
    _escribir_hoja_resumen(wb, causas)

    # Poner Resumen al principio (moverlo al índice 0)
    wb.move_sheet("Resumen", offset=-len(wb.sheetnames) + 1)

    # Guardar — si el archivo del día ya existe, agregar sufijo incremental _2, _3...
    os.makedirs(REPORTES_DIR, exist_ok=True)
    fecha = datetime.date.today().isoformat()
    ruta  = os.path.join(REPORTES_DIR, f"Reporte_{fecha}.xlsx")
    if os.path.exists(ruta):
        contador = 2
        while os.path.exists(os.path.join(REPORTES_DIR, f"Reporte_{fecha}_{contador}.xlsx")):
            contador += 1
        ruta = os.path.join(REPORTES_DIR, f"Reporte_{fecha}_{contador}.xlsx")
    wb.save(ruta)

    log.info(f"Reporte guardado: {ruta}")

    # Resumen a consola
    por_clas = {}
    for c in causas:
        k = c.get("_clasificacion", "SIN PDF")
        por_clas[k] = por_clas.get(k, 0) + 1

    log.info("=" * 55)
    log.info(f"Reporte: {os.path.basename(ruta)}")
    for clas in ("CON DEUDA EXTRAÍDA", "SIN PDF", "SIN MONTO EN PDF"):
        log.info(f"  {clas:<20}: {por_clas.get(clas, 0)}")
    log.info("=" * 55)

    return ruta


# ─────────────────────────────────────────────────────────────────
# Standalone: genera reporte con datos sintéticos para probar formato
# ─────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    import sys

    if "--demo" in sys.argv:
        # Genera un reporte de demostración sin correr el pipeline
        import random
        random.seed(42)

        comunas_r  = ["Valparaíso", "Concepción", "Temuco", "La Serena", "Iquique"]
        tipos_proc = ["ejecutivo", "ley_bancos"]
        tipos_doc  = {"ejecutivo": "mandamiento", "ley_bancos": "bases_remate"}

        causas_demo = []
        for i in range(1, 26):
            proc  = random.choice(tipos_proc)
            deuda = random.randint(20_000_000, 150_000_000)
            desc  = random.random() > 0.2

            causas_demo.append({
                "rol":                str(30000 + i),
                "año":                str(random.randint(2015, 2023)),
                "corte":              f"C.A. de {random.choice(comunas_r)}",
                "tribunal":           f"{i}° Juzgado Civil de Concepción",
                "demandante":         random.choice(["Banco BCI", "Banco Itaú Chile", "Banco Santander"]),
                "direccion":          f"Calle Demo {i * 100}",
                "comuna":             random.choice(comunas_r),
                "region_rm":          False,
                "tipo_procedimiento": proc,
                "tipo_documento":     tipos_doc[proc] if desc else "",
                "descargado":         desc,
                "ruta_pdf":           "",
                "monto_deuda_clp":    deuda if desc else 0,
                "monto_original":     f"${deuda:,}" if desc else "",
            })

        print(f"Generando reporte demo con {len(causas_demo)} causas...")
        ruta = generar_reporte(causas_demo)
        print(f"Reporte creado: {ruta}")

    else:
        # Pipeline completo M1 → M5
        print("Modo standalone v1 deshabilitado (Tanda D). Usa los flags "
              "del módulo o el pipeline completo: python main.py --docx ...")
        raise SystemExit(0)
