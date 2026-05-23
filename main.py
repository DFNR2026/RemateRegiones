"""
main.py — Orquestador del Sistema de Análisis de Remates Judiciales

Encadena los 5 módulos en secuencia e imprime un resumen de cada etapa.

Uso:
    python main.py              # pipeline completo (M1 → M5)
    python main.py --demo       # 25 causas sintéticas (omite M1 y M2)
    python main.py --sin-ojv    # omite M2 (usa PDFs ya descargados en Descargas/)
    python main.py --hasta 3    # detiene el pipeline después del Módulo N
    python main.py --silencio   # suprime logs de módulos (solo muestra resúmenes)
    python main.py --docx "RESUMEN.docx"          # DOCX semanal (v2 Regex+Haiku por defecto)
    python main.py --v1 --docx "RESUMEN.docx"     # DOCX semanal forzando Sonnet (v1)

# ===========================================================================
# ANTES DE CADA CORRIDA DE PRUEBA — limpiar estado previo:
#
#   python limpiar_cache.py
#
# Esto borra:
#   1. Todo el contenido de D:\\Remates\\Descargas\\
#   2. Las filas de datos de la hoja CAUSAS en causas_ojv.xlsx
#      (deja solo el encabezado, para que M1 no filtre causas ya procesadas)
#
# Si no se limpia entre pruebas:
#   - M1 marcará las causas como "ya en historial" y devolverá 0 causas nuevas
#   - Los PDFs en Descargas/ pueden contener mandamientos de causas incorrectas
#     (bug de "primera fila" corregido en ojv_remates.py v10.1)
# ===========================================================================
"""

import os
import sys
import time
import shutil
import argparse
import logging
from datetime import datetime

# ─────────────────────────────────────────────────────────────────
# Dual-log: consola + archivo en logs/
# ─────────────────────────────────────────────────────────────────
_LOGS_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logs")
os.makedirs(_LOGS_DIR, exist_ok=True)
_LOG_FILE = os.path.join(
    _LOGS_DIR,
    f"ejecucion_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log",
)

_root = logging.getLogger()
for _h in _root.handlers[:]:
    _root.removeHandler(_h)
_root.setLevel(logging.INFO)

_FMT = logging.Formatter("%(asctime)s %(message)s")

_console_handler = logging.StreamHandler(sys.stdout)
_console_handler.setFormatter(_FMT)
_root.addHandler(_console_handler)

_file_handler = logging.FileHandler(_LOG_FILE, encoding="utf-8")
_file_handler.setFormatter(_FMT)
_root.addHandler(_file_handler)

log = logging.getLogger("main")


class _TeeWriter:
    """Redirige print() a consola + archivo de log simultáneamente."""

    def __init__(self, original_stream, log_file_handle):
        self._orig = original_stream
        self._log = log_file_handle

    def write(self, text):
        self._orig.write(text)
        self._orig.flush()
        self._log.write(text)
        self._log.flush()

    def flush(self):
        self._orig.flush()
        self._log.flush()

    # Necesario para que logging no falle al verificar el stream
    def fileno(self):
        return self._orig.fileno()


_log_fh = open(_LOG_FILE, "a", encoding="utf-8")
sys.stdout = _TeeWriter(sys.__stdout__, _log_fh)
sys.stderr = _TeeWriter(sys.__stderr__, _log_fh)

# ─────────────────────────────────────────────────────────────────
# Imports de módulos del proyecto
# ─────────────────────────────────────────────────────────────────
from modulo1_parser       import parsear_diarios, parsear_docx_semanal

from modulo2_ojv          import procesar_causas_ojv
from modulo3_extractor    import extraer_montos
from modulo5_reporte      import generar_reporte, actualizar_historial


# ─────────────────────────────────────────────────────────────────
# Utilidades de presentación
# ─────────────────────────────────────────────────────────────────

_W = 60   # ancho de línea de separador


def _sep(titulo: str = "") -> None:
    print()
    if titulo:
        print("=" * _W)
        print(f"  {titulo}")
    print("=" * _W)


def _ok(modulo: int, msg: str, tiempo_s: float) -> None:
    print(f"  [M{modulo}] {msg}  ({tiempo_s:.1f}s)")


def _resumen_final(causas: list[dict], elapsed_s: float, ruta_reporte: str) -> None:
    """Imprime el resumen ejecutivo al terminar el pipeline."""
    total = len(causas)
    desc  = sum(1 for c in causas if c.get("descargado"))
    monts = sum(1 for c in causas if c.get("monto_deuda_clp"))
    por_clas = {}
    for c in causas:
        k = c.get("_clasificacion", "SIN PDF")
        por_clas[k] = por_clas.get(k, 0) + 1

    _sep("RESUMEN EJECUTIVO")
    mins, segs = divmod(int(elapsed_s), 60)
    print(f"  Tiempo total           : {mins}m {segs}s")
    print(f"  Causas procesadas      : {total}")
    print(f"  Documentos descargados : {desc}")
    print(f"  Montos extraídos       : {monts}")
    print()
    print("  ESTADO DE DEUDA            N")
    print("  " + "-" * 30)
    for clas in ("CON DEUDA EXTRAÍDA", "SIN PDF", "SIN MONTO EN PDF"):
        n = por_clas.get(clas, 0)
        barra = "#" * n
        print(f"  {clas:<22} {n:>3}  {barra}")
    print()
    if ruta_reporte:
        print(f"  Reporte: {ruta_reporte}")
    _sep()


# ─────────────────────────────────────────────────────────────────
# Causas demo (para pruebas sin M1/M2)
# ─────────────────────────────────────────────────────────────────

def _causas_demo() -> list[dict]:
    """Genera 25 causas sintéticas para probar M3→M5 sin correr M1/M2."""
    import random, os
    from config import DESCARGAS_DIR
    random.seed(42)

    comunas_rm  = ["Maipú", "La Florida", "Santiago", "Puente Alto", "San Bernardo",
                   "Peñalolén", "Las Condes", "Ñuñoa", "San Miguel", "Lo Espejo"]
    comunas_reg = ["Valparaíso", "Concepción", "Temuco", "La Serena",
                   "Iquique", "Antofagasta", "Rancagua", "Talca"]
    bancos = ["Banco BCI", "Banco Itaú Chile", "Banco Santander", "Banco de Chile"]

    causas = []
    for i in range(1, 26):
        es_rm   = i <= 15
        proc    = "ejecutivo" if random.random() > 0.2 else "ley_bancos"
        tipo_doc = "mandamiento" if proc == "ejecutivo" else "bases_remate"
        desc    = random.random() > 0.15
        deuda   = random.randint(15_000_000, 120_000_000) if desc else 0
        comuna  = random.choice(comunas_rm if es_rm else comunas_reg)

        # Simular ruta_pdf (el archivo no existe, pero M3 lo manejará)
        etq     = f"C-{30000 + i}-{random.randint(2015, 2023)}"
        sufijo  = "MANDAMIENTO" if proc == "ejecutivo" else "BASES_REMATE"
        ruta    = os.path.join(DESCARGAS_DIR, f"{etq}_{sufijo}.pdf")

        causas.append({
            "rol":                str(30000 + i),
            "año":                str(random.randint(2015, 2023)),
            "corte":              "C.A. de Santiago" if es_rm
                                  else f"C.A. de {random.choice(comunas_reg)}",
            "tribunal":           f"{i}° Juzgado Civil de {'Santiago' if es_rm else comuna}",
            "demandante":         random.choice(bancos),
            "direccion":          f"Calle Demo {i * 100}",
            "comuna":             comuna,
            "region_rm":          es_rm,
            "tipo_procedimiento": proc,
            "tipo_documento":     tipo_doc if desc else "",
            "descargado":         desc,
            "ruta_pdf":           ruta if desc else "",
            "monto_deuda_clp":    deuda,
            "monto_original":     f"${deuda:,}" if deuda else "",
        })
    return causas


# ─────────────────────────────────────────────────────────────────
# Pipeline principal
# ─────────────────────────────────────────────────────────────────

def main() -> None:
    parser = argparse.ArgumentParser(
        description="Sistema de Análisis de Remates Judiciales — Chile",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Ejemplos:
  python main.py                        # pipeline completo
  python main.py --demo                 # prueba con datos sintéticos (sin M1/M2)
  python main.py --sin-ojv              # omite M2 (usa PDFs ya descargados)
  python main.py --hasta 3              # detiene tras Módulo 3
  python main.py --silencio             # solo resúmenes, sin logs de módulos
  python main.py --limpiar-historial    # M1 ignora historial CAUSAS (testing)
  python main.py --diarios "D:\\Remates\\Diarios test"  # carpeta alternativa de PDFs
  python main.py --docx "D:\\Remates\\RESUMEN.docx"     # DOCX semanal (v2 por defecto)
  python main.py --v1 --docx "D:\\Remates\\RESUMEN.docx"  # DOCX forzando Sonnet (v1)
        """,
    )
    parser.add_argument("--demo",     action="store_true",
                        help="Usar 25 causas sintéticas (omite M1 y M2)")
    parser.add_argument("--sin-ojv",  action="store_true",
                        help="Omitir M2 (usar PDFs ya descargados en Descargas/)")
    parser.add_argument("--hasta",    type=int, default=5, metavar="N",
                        help="Detener el pipeline después del Módulo N (1-5)")
    parser.add_argument("--silencio", action="store_true",
                        help="Suprimir logs de módulos (mostrar solo resúmenes)")
    parser.add_argument("--limpiar-historial", action="store_true",
                        help="M1 ignora la hoja CAUSAS (trata historial como vacío). "
                             "El Excel NO se modifica. Solo para testing.")
    parser.add_argument("--diarios",  type=str, default=None, metavar="RUTA",
                        help="Carpeta alternativa de PDFs de entrada (default: Diarios/)")
    parser.add_argument("--docx",     type=str, default=None, metavar="RUTA",
                        help="Ruta al DOCX semanal consolidado (reemplaza PDFs diarios)")
    parser.add_argument("--docx-dir", type=str, default=None, metavar="RUTA",
                        help="Carpeta con varios DOCX semanales. Procesa todos los .docx "
                             "que contenga y los elimina al terminar el pipeline.")
    parser.add_argument("--v1",       action="store_true",
                        help="Forzar parser v1 (Sonnet) para DOCX. "
                             "Por defecto --docx usa v2 (Regex+Haiku, ~60x mas barato).")
    args = parser.parse_args()

    if args.docx and args.docx_dir:
        parser.error("--docx y --docx-dir son mutuamente excluyentes")

    if args.silencio:
        # Suprimir todos los loggers excepto el de main
        logging.root.setLevel(logging.WARNING)
        log.setLevel(logging.INFO)

    t_total = time.time()
    ruta_reporte = ""
    causas: list[dict] = []
    docx_procesados_paths: list[str] = []  # DOCX a borrar tras pipeline completo (modo --docx-dir)

    _sep("SISTEMA DE ANÁLISIS DE REMATES JUDICIALES")
    print(f"  Log: {_LOG_FILE}")

    try:
        # ── Modo demo: saltar M1 y M2 ──────────────────────────
        if args.demo:
            causas = _causas_demo()
            print(f"  [DEMO] {len(causas)} causas sintéticas generadas")
            modulo_inicio = 3
        else:
            modulo_inicio = 1

        # ── Módulo 1: Parser PDFs  o  DOCX semanal ──────────────────
        if modulo_inicio <= 1 <= args.hasta:
            if args.docx_dir:
                # ── Modo carpeta: procesar TODOS los .docx de la carpeta ──
                if args.v1:
                    _sep("MÓDULO 1 v1 — Parsear carpeta de DOCX (Sonnet)")
                    _parser_docx = parsear_docx_semanal
                    _tag = "M1"
                else:
                    from v2_experimental.modulo1_v2 import parsear_docx_v2
                    _sep("MÓDULO 1 v2 — Parsear carpeta de DOCX (Regex+Haiku)")
                    _parser_docx = parsear_docx_v2
                    _tag = "M1v2"
                print(f"  [{_tag}] Carpeta DOCX: {args.docx_dir}")
                if args.limpiar_historial:
                    print(f"  [{_tag}] Modo testing: historial CAUSAS ignorado (Excel intacto)")

                # Listar .docx (case-insensitive), ignorar temporales de Word (~$)
                docx_files = sorted(
                    f for f in os.listdir(args.docx_dir)
                    if f.lower().endswith(".docx") and not f.startswith("~$")
                )
                if not docx_files:
                    print(f"  [{_tag}] Sin DOCX en {args.docx_dir}")
                    return

                t = time.time()
                causas = []
                roles_vistos = set()
                for _nombre in docx_files:
                    _ruta = os.path.join(args.docx_dir, _nombre)
                    _nuevas = _parser_docx(_ruta, ignorar_historial=args.limpiar_historial)
                    # Dedup dentro del lote por (ROL, AÑO): el historial persistido
                    # no se actualiza hasta que corre M2+, asi que dos DOCX del mismo
                    # lote podrian traer el mismo ROL.
                    _filtradas = []
                    for _c in _nuevas:
                        _clave = (str(_c.get("rol", "")).strip(),
                                  str(_c.get("año", "")).strip())
                        if _clave in roles_vistos:
                            continue
                        roles_vistos.add(_clave)
                        _filtradas.append(_c)
                    causas.extend(_filtradas)
                    print(f"  [{_tag}] Procesando: {_nombre} -> {len(_filtradas)} causas nuevas")

                print(f"  [{_tag}] Total acumulado: {len(causas)} causas nuevas de {len(docx_files)} DOCX")
                _ok(1, f"{len(causas)} causas nuevas ({len(docx_files)} DOCX)", time.time() - t)

                # Registrar rutas para borrarlas SOLO al final del pipeline (M1->M5).
                # No se borran aca: si M2/M3/M5 fallan, los DOCX quedan para reintentar.
                docx_procesados_paths = [os.path.join(args.docx_dir, n) for n in docx_files]

            elif args.docx and args.v1:
                # Forzar v1 (Sonnet) para DOCX
                _sep("MÓDULO 1 v1 — Parsear DOCX semanal (Sonnet)")
                print(f"  [M1] Archivo DOCX: {args.docx}")
                if args.limpiar_historial:
                    print("  [M1] Modo testing: historial CAUSAS ignorado (Excel intacto)")
                t = time.time()
                causas = parsear_docx_semanal(
                    args.docx,
                    ignorar_historial=args.limpiar_historial,
                )
                _ok(1, f"{len(causas)} causas nuevas (v1 Sonnet)", time.time() - t)
            elif args.docx:
                # Default: v2 (Regex+Haiku) para DOCX
                from v2_experimental.modulo1_v2 import parsear_docx_v2
                _sep("MÓDULO 1 v2 — Parsear DOCX semanal (Regex+Haiku)")
                print(f"  [M1v2] Archivo DOCX: {args.docx}")
                if args.limpiar_historial:
                    print("  [M1v2] Modo testing: historial CAUSAS ignorado (Excel intacto)")
                t = time.time()
                causas = parsear_docx_v2(
                    args.docx,
                    ignorar_historial=args.limpiar_historial,
                )
                _ok(1, f"{len(causas)} causas nuevas (v2 Regex+Haiku)", time.time() - t)
            else:
                _sep("MÓDULO 1 — Parsear PDFs del Diario P&L")
                if args.limpiar_historial:
                    print("  [M1] Modo testing: historial CAUSAS ignorado (Excel intacto)")
                if args.diarios:
                    print(f"  [M1] Carpeta de entrada: {args.diarios}")
                t = time.time()
                m1_kwargs = {"ignorar_historial": args.limpiar_historial}
                if args.diarios:
                    m1_kwargs["directorio"] = args.diarios
                causas = parsear_diarios(**m1_kwargs)
                _ok(1, f"{len(causas)} causas nuevas detectadas", time.time() - t)

            if not causas:
                print()
                print("  Sin causas nuevas. Todos los registros ya estaban en el historial.")
                return

            # Mover PDFs procesados a Diarios_Procesados/
            # (solo si se usó la carpeta por defecto; carpeta custom y DOCX no se tocan)
            if not args.diarios and not args.docx and not args.docx_dir:
                from config import DIARIOS_DIR, DIARIOS_PROCESADOS_DIR
                os.makedirs(DIARIOS_PROCESADOS_DIR, exist_ok=True)
                pdfs_movidos = 0
                for f in os.listdir(DIARIOS_DIR):
                    if f.lower().endswith(".pdf"):
                        src = os.path.join(DIARIOS_DIR, f)
                        dst = os.path.join(DIARIOS_PROCESADOS_DIR, f)
                        shutil.move(src, dst)
                        pdfs_movidos += 1
                if pdfs_movidos:
                    print(f"  [M1] {pdfs_movidos} PDF(s) movidos a Diarios_Procesados/")

        # ── Módulo 2: OJV + Descarga ────────────────────────────
        if modulo_inicio <= 2 <= args.hasta and not args.sin_ojv:
            _sep("MÓDULO 2 — Consulta OJV y descarga de documentos")
            t = time.time()
            causas = procesar_causas_ojv(causas, n_workers=5)
            desc = sum(1 for c in causas if c.get("descargado"))
            _ok(2, f"{desc}/{len(causas)} documentos descargados", time.time() - t)
        elif args.sin_ojv:
            print("  [M2] Omitido (--sin-ojv). Usando PDFs existentes en Descargas/")
            from config import DESCARGAS_DIR

            # Indexar PDFs disponibles en Descargas/ por nombre de archivo
            _pdfs_disponibles = {
                f: os.path.join(DESCARGAS_DIR, f)
                for f in os.listdir(DESCARGAS_DIR)
                if f.lower().endswith(".pdf")
            }

            encontrados = 0
            for c in causas:
                etq = f"C-{c['rol']}-{c['año']}"
                c.setdefault("tipo_procedimiento", "")
                # Buscar mandamiento o bases de remate para esta causa
                for sufijo, tipo_doc in [("_MANDAMIENTO.pdf",  "mandamiento"),
                                         ("_BASES_REMATE.pdf", "bases_remate")]:
                    nombre = f"{etq}{sufijo}"
                    if nombre in _pdfs_disponibles:
                        c["descargado"]   = True
                        c["ruta_pdf"]     = _pdfs_disponibles[nombre]
                        c["tipo_documento"] = tipo_doc
                        encontrados += 1
                        break
                else:
                    c.setdefault("descargado",     False)
                    c.setdefault("ruta_pdf",       "")
                    c.setdefault("tipo_documento", "")

            print(f"  [M2] {encontrados}/{len(causas)} PDFs detectados en Descargas/")

        if args.hasta < 3:
            _resumen_final(causas, time.time() - t_total, ruta_reporte)
            return

        # ── Módulo 3: Extracción de montos ─────────────────────
        _sep("MÓDULO 3 — Extracción de montos de deuda")
        t = time.time()
        causas = extraer_montos(causas)
        monts = sum(1 for c in causas if c.get("monto_deuda_clp"))
        _ok(3, f"{monts}/{len(causas)} montos extraídos", time.time() - t)

        if args.hasta < 4:
            _resumen_final(causas, time.time() - t_total, ruta_reporte)
            return

        # ── Módulo 5: Reporte ───────────────────────────────────
        _sep("MÓDULO 5 — Reporte y actualización de historial")
        t = time.time()
        actualizar_historial(causas)
        ruta_reporte = generar_reporte(causas)
        _ok(5, f"Reporte generado: {ruta_reporte}", time.time() - t)

        # Borrar los DOCX procesados SOLO tras completar el pipeline (M1->M5).
        # Si un modulo posterior a M1 hubiera fallado, una excepcion habria
        # saltado este bloque y los DOCX quedarian para reintentar.
        # Solo se borran los archivos que se procesaron en esta corrida.
        if docx_procesados_paths:
            _dtag = "M1" if args.v1 else "M1v2"
            _docx_borrados = 0
            for _p in docx_procesados_paths:
                try:
                    if os.path.exists(_p):
                        os.remove(_p)
                        _docx_borrados += 1
                except OSError as _e:
                    print(f"  [{_dtag}] No se pudo borrar {os.path.basename(_p)}: {_e}")
            if _docx_borrados:
                print(f"  [{_dtag}] {_docx_borrados} DOCX procesados eliminados de la carpeta")

    except KeyboardInterrupt:
        print()
        print("  Detenido por el usuario (Ctrl+C).")
        if causas:
            print(f"  {len(causas)} causas en memoria — generando reporte parcial...")
            try:
                ruta_reporte = generar_reporte(causas)
                print(f"  Reporte parcial: {ruta_reporte}")
            except Exception as e:
                print(f"  Error al generar reporte parcial: {e}")

    except Exception as e:
        print()
        print(f"  ERROR FATAL: {e}")
        import traceback
        traceback.print_exc()
        if causas:
            print(f"  {len(causas)} causas en memoria hasta el punto de fallo.")
        sys.exit(1)

    # ── Resumen final ───────────────────────────────────────────
    _resumen_final(causas, time.time() - t_total, ruta_reporte)


if __name__ == "__main__":
    main()
