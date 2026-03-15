"""
Módulo 3: Extractor de Montos de Deuda

Input:  lista de dicts del Módulo 2 (con campos: ruta_pdf, tipo_documento,
        tipo_procedimiento, descargado)
Output: misma lista enriquecida con:
        - monto_deuda_clp : int | None  (monto en pesos chilenos; None si no se pudo extraer)
        - monto_original  : str         (texto literal extraído del PDF; "" si no)

Lógica:
  - Mandamientos  → busca "capital adeudado", "pague la suma de", etc.
  - Bases remate  → busca "mínimo para las posturas", "precio mínimo", etc.
  - Montos en UF → convertidos a CLP con el valor UF del día (mindicador.cl)
"""

import os
import re
import logging

import sys
import contextlib

import fitz          # PyMuPDF
fitz.TOOLS.mupdf_warnings(False)          # silenciar warnings cosméticos
fitz.TOOLS.mupdf_display_errors(False)    # silenciar errores de annotations/rendering
import requests


@contextlib.contextmanager
def _silenciar_stderr():
    """Redirige stderr a devnull temporalmente (atrapa mensajes del core C de MuPDF)."""
    old = sys.stderr
    try:
        sys.stderr = open(os.devnull, "w")
        yield
    finally:
        sys.stderr.close()
        sys.stderr = old

from config import DESCARGAS_DIR

# ─────────────────────────────────────────────────────────────────
# OCR fallback (pytesseract + Tesseract)
# Se activa solo cuando PyMuPDF extrae menos de 100 caracteres
# (PDF escaneado como imagen en vez de texto extraíble).
# ─────────────────────────────────────────────────────────────────

_TESSERACT_CMD  = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
_TESSDATA_DIR   = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".tessdata")
_OCR_DISPONIBLE = False

try:
    import pytesseract
    from PIL import Image as _PILImage
    pytesseract.pytesseract.tesseract_cmd = _TESSERACT_CMD
    os.environ.setdefault("TESSDATA_PREFIX", _TESSDATA_DIR)
    pytesseract.get_tesseract_version()   # lanza excepción si el binario no existe
    _OCR_DISPONIBLE = True
except Exception:
    pass  # OCR no disponible; se logeará al usarlo

logging.basicConfig(level=logging.INFO, format="%(asctime)s [M3] %(message)s")
log = logging.getLogger(__name__)


# ─────────────────────────────────────────────────────────────────
# Valor UF
# ─────────────────────────────────────────────────────────────────

_uf_cache: dict = {}

UF_FALLBACK = 38_500.0   # actualizar si la API falla y el valor es muy distante


def obtener_uf_hoy() -> float:
    """
    Obtiene el valor actual de la UF desde la API gratuita de mindicador.cl.
    Guarda el resultado en caché para no llamar la API múltiples veces.
    """
    if "v" in _uf_cache:
        return _uf_cache["v"]
    try:
        resp = requests.get(
            "https://mindicador.cl/api/uf",
            timeout=10,
            headers={"User-Agent": "Mozilla/5.0"},
        )
        resp.raise_for_status()
        valor = float(resp.json()["serie"][0]["valor"])
        log.info(f"UF actual: ${valor:,.2f}")
        _uf_cache["v"] = valor
        return valor
    except Exception as e:
        log.warning(f"API mindicador.cl falló ({e}). Usando fallback ${UF_FALLBACK:,.0f}")
        _uf_cache["v"] = UF_FALLBACK
        return UF_FALLBACK


# ─────────────────────────────────────────────────────────────────
# Regex y patrones
# ─────────────────────────────────────────────────────────────────

# Número en formato UF chileno: "1.767,802476"  (punto=miles, coma=decimal)
# Captura: grupos con punto de miles + coma decimal, o número simple con coma
_RE_NUM_UF = re.compile(r'\d{1,3}(?:\.\d{3})*,\d+|\d+,\d+')

# Sufijo que indica que el número anterior es UF
_RE_UF_SUFIJO = re.compile(r'U\.?F\.?|Unidades?\s+de\s+Fomento', re.IGNORECASE)

# Número en formato CLP chileno precedido por $: "$15.234.567"
# Captura solo el número (sin el $)
_RE_NUM_CLP = re.compile(r'\$\s*(\d{1,3}(?:\.\d{3})+(?:,\d+)?)')


# ── Contextos de búsqueda para MANDAMIENTOS ──
# En orden de especificidad (los más específicos primero)
_CTX_MANDAMIENTO = [
    # Orden de pago directa: "pague/n la suma de"
    r'pague[n]?\s+la\s+suma\s+de\s*',
    r'pague[n]?\s+(?:la\s+)?cantidad\s+de\s*',
    # Condena a pagar
    r'condena(?:do|ndo)?\s+[^\n]{0,80}?pagar\s+la\s+suma\s+de\s*',
    r'deber[áa]\s+pagar\s+la\s+suma\s+de\s*',
    r'obliga(?:do)?\s+[^\n]{0,60}?pagar\s+la\s+suma\s+de\s*',
    # Mención directa de capital
    r'capital\s+adeudado\s+(?:de\s+)?',
    r'capital\s+insoluto\s+(?:de\s+)?',
    r'por\s+concepto\s+de\s+capital[^.]{0,50}?suma\s+de\s*',
    r'monto\s+(?:total\s+)?(?:de\s+la\s+deuda\s+)?de\s*',
    # Genérico (menor precisión, usar al final)
    r'la\s+suma\s+de\s*',
    r'la\s+cantidad\s+de\s*',
]

# ── Contextos de búsqueda para BASES DE REMATE ──
# Cada patrón acepta tanto formas con acento (mínimo, será) como sin acento (minimo, sera)
_CTX_BASES = [
    # Fórmula clásica del otrosí
    r'm[íi]nimo\s+para\s+las?\s+posturas\s+(?:ser[áa][n]?|se\s+fij[ao][n]?)\s*(?:en|de)?\s*',
    r'm[íi]nimo\s+para\s+las?\s+posturas\s*[=:]\s*',
    r'm[íi]nimo\s+para\s+las?\s+posturas[^.]{0,40}?(?:de\s+)?',
    # Mínimo de la subasta
    r'm[íi]nimo\s+de\s+la\s+subasta\s+(?:ser[áa][n]?|es)\s*(?:de\s+)?',
    r'm[íi]nimo\s+de\s+la\s+subasta\s*[=:]\s*',
    # Precio mínimo
    r'precio\s+m[íi]nimo\s+(?:de\s+(?:la\s+)?(?:subasta\s+)?)?(?:ser[áa][n]?|es)?\s*(?:de\s+)?',
    r'precio\s+m[íi]nimo\s*[=:]\s*',
    # Fallback genérico
    r'm[íi]nimo\s+de\s*',
]

# Ventana de búsqueda (caracteres a analizar tras el contexto)
_VENTANA = 140


# ─────────────────────────────────────────────────────────────────
# Extracción de texto del PDF
# ─────────────────────────────────────────────────────────────────

_OCR_UMBRAL = 100   # chars mínimos para considerar que PyMuPDF extrajo texto útil


def _ocr_pdf(ruta_pdf: str) -> str:
    """
    Extrae texto de un PDF escaneado usando OCR (pytesseract).
    Renderiza cada página a 300 dpi con PyMuPDF y la pasa a Tesseract (spa+eng).
    Retorna "" si OCR no está disponible o falla.
    """
    if not _OCR_DISPONIBLE:
        log.warning("OCR no disponible (pytesseract/Tesseract no instalado)")
        return ""
    try:
        with _silenciar_stderr():
            doc = fitz.open(ruta_pdf)
            paginas_texto = []
            matriz = fitz.Matrix(300 / 72, 300 / 72)   # escala a 300 dpi
            for page in doc:
                pix = page.get_pixmap(matrix=matriz, colorspace=fitz.csGRAY)
                img = _PILImage.frombytes("L", (pix.width, pix.height), pix.samples)
                texto_pag = pytesseract.image_to_string(img, lang="spa+eng", config="--psm 6")
                paginas_texto.append(texto_pag)
            doc.close()
        texto = "\n".join(paginas_texto)
        texto = texto.replace("\xad", "")
        texto = re.sub(r" {2,}", " ", texto)
        return texto
    except Exception as e:
        log.error(f"OCR falló en {ruta_pdf}: {e}")
        return ""


def _extraer_texto_pdf(ruta_pdf: str) -> str:
    """
    Extrae texto del PDF. Intenta PyMuPDF primero; si devuelve menos de
    _OCR_UMBRAL caracteres (PDF escaneado), intenta OCR con pytesseract.
    """
    try:
        with _silenciar_stderr():
            doc = fitz.open(ruta_pdf)
            paginas = [page.get_text("text") for page in doc]
            doc.close()
        texto = "\n".join(paginas)
        texto = texto.replace("\xad", "")
        texto = re.sub(r" {2,}", " ", texto)
    except Exception as e:
        log.error(f"Error leyendo {ruta_pdf}: {e}")
        return ""

    if len(texto.strip()) < _OCR_UMBRAL:
        log.info(f"  PDF escaneado ({len(texto.strip())} chars) — intentando OCR...")
        texto_ocr = _ocr_pdf(ruta_pdf)
        if len(texto_ocr.strip()) > len(texto.strip()):
            log.info(f"  OCR exitoso: {len(texto_ocr.strip())} chars extraídos")
            return texto_ocr
        log.warning(f"  OCR sin mejora ({len(texto_ocr.strip())} chars)")

    return texto


# ─────────────────────────────────────────────────────────────────
# Motor de extracción
# ─────────────────────────────────────────────────────────────────

def _parsear_uf(s: str) -> float:
    """'1.767,802476' → 1767.802476"""
    return float(s.replace(".", "").replace(",", "."))


def _parsear_clp(s: str) -> int:
    """'15.234.567' → 15234567"""
    return int(s.replace(".", "").replace(",", ""))


def _extraer_monto_con_contextos(
    texto: str, contextos: list[str]
) -> tuple[str, float, str]:
    """
    Busca en texto uno de los contextos y extrae el monto (UF o CLP) que le sigue.

    Returns:
        (texto_original, valor_numerico, moneda)  donde moneda = "UF" | "CLP"
        ("", 0.0, "")  si no se encontró nada
    """
    for ctx in contextos:
        patron = re.compile(ctx, re.IGNORECASE | re.DOTALL)
        for m in patron.finditer(texto):
            fragmento = texto[m.end() : m.end() + _VENTANA]
            # Normalizar espacios internos (saltos de línea entre número y sufijo)
            frag = re.sub(r"\s+", " ", fragmento)

            # ── Intentar UF primero ──
            mu = _RE_NUM_UF.search(frag)
            if mu:
                pos = mu.end()
                resto = frag[pos : pos + 35]
                if _RE_UF_SUFIJO.search(resto):
                    try:
                        valor = _parsear_uf(mu.group())
                        if valor > 0.5:  # mínimo razonable: 0.5 UF
                            return (mu.group() + " UF", valor, "UF")
                    except ValueError:
                        pass

            # ── Intentar CLP ──
            mc = _RE_NUM_CLP.search(frag)
            if mc:
                try:
                    valor = _parsear_clp(mc.group(1))
                    if valor >= 100_000:  # mínimo razonable: $100.000
                        return ("$" + mc.group(1), float(valor), "CLP")
                except ValueError:
                    pass

    return ("", 0.0, "")


def _buscar_monto_amplio(texto: str) -> tuple[str, float, str]:
    """
    Último recurso: busca cualquier monto UF o CLP en el documento completo.
    Toma el primero que aparezca dentro de rangos razonables para una deuda.
    Sólo se usa si los contextos específicos fallaron.
    """
    # Intentar UF (rango: 1 UF a 200.000 UF ≈ $7.700M)
    for m in _RE_NUM_UF.finditer(texto):
        pos = m.end()
        resto = re.sub(r"\s+", " ", texto[pos : pos + 35])
        if _RE_UF_SUFIJO.search(resto):
            try:
                valor = _parsear_uf(m.group())
                if 1.0 <= valor <= 200_000:
                    return (m.group() + " UF", valor, "UF")
            except ValueError:
                pass

    # Intentar CLP (rango: $1.000.000 a $5.000.000.000)
    for m in _RE_NUM_CLP.finditer(texto):
        try:
            valor = _parsear_clp(m.group(1))
            if 1_000_000 <= valor <= 5_000_000_000:
                return ("$" + m.group(1), float(valor), "CLP")
        except ValueError:
            pass

    return ("", 0.0, "")


# ─────────────────────────────────────────────────────────────────
# Extractores por tipo de documento
# ─────────────────────────────────────────────────────────────────

def _extraer_de_mandamiento(texto: str) -> tuple[str, float, str]:
    orig, valor, moneda = _extraer_monto_con_contextos(texto, _CTX_MANDAMIENTO)
    if not orig:
        orig, valor, moneda = _buscar_monto_amplio(texto)
        if orig:
            log.debug("  Monto obtenido por búsqueda amplia (mandamiento)")
    return orig, valor, moneda


def _extraer_de_bases_remate(texto: str) -> tuple[str, float, str]:
    orig, valor, moneda = _extraer_monto_con_contextos(texto, _CTX_BASES)
    if not orig:
        orig, valor, moneda = _buscar_monto_amplio(texto)
        if orig:
            log.debug("  Monto obtenido por búsqueda amplia (bases remate)")
    return orig, valor, moneda


# ─────────────────────────────────────────────────────────────────
# FUNCIÓN PÚBLICA — interface para el orquestador
# ─────────────────────────────────────────────────────────────────

def extraer_montos(causas: list[dict]) -> list[dict]:
    """
    Recibe la lista de causas del Módulo 2 y para cada causa con PDF
    descargado extrae el monto de deuda y lo convierte a CLP.

    Args:
        causas: lista de dicts (output del Módulo 2)

    Returns:
        Misma lista enriquecida con: monto_deuda_clp (int), monto_original (str)
    """
    log.info(f"Iniciando Módulo 3 — {len(causas)} causa(s)")

    uf_hoy = obtener_uf_hoy()

    extraidos = 0
    sin_pdf   = 0
    fallidos  = 0

    for causa in causas:
        causa.setdefault("monto_deuda_clp", None)
        causa.setdefault("monto_original", "")

        ruta_pdf = causa.get("ruta_pdf", "")
        etiqueta = f"C-{causa['rol']}-{causa['año']}"

        # Sin PDF → saltar
        if not ruta_pdf or not causa.get("descargado"):
            sin_pdf += 1
            causa.setdefault("motivo_fallo", "M3: PDF no descargado")
            continue

        if not os.path.exists(ruta_pdf):
            log.warning(f"  {etiqueta}: ruta_pdf no existe en disco ({ruta_pdf})")
            fallidos += 1
            causa.setdefault("motivo_fallo", "M3: PDF no descargado")
            continue

        texto = _extraer_texto_pdf(ruta_pdf)
        if not texto.strip():
            log.warning(f"  {etiqueta}: PDF vacío o ilegible")
            fallidos += 1
            causa.setdefault("motivo_fallo", "M3: monto no extraído")
            continue

        tipo_doc = causa.get("tipo_documento", "")

        if tipo_doc == "mandamiento":
            orig, valor, moneda = _extraer_de_mandamiento(texto)
        elif tipo_doc == "bases_remate":
            orig, valor, moneda = _extraer_de_bases_remate(texto)
        else:
            log.warning(f"  {etiqueta}: tipo_documento desconocido '{tipo_doc}'")
            fallidos += 1
            causa.setdefault("motivo_fallo", "M3: monto no extraído")
            continue

        if not orig:
            log.warning(f"  {etiqueta}: no se encontró monto en {tipo_doc}")
            fallidos += 1
            causa.setdefault("motivo_fallo", "M3: monto no extraído")
            continue

        # Convertir a CLP
        if moneda == "UF":
            monto_clp = int(round(valor * uf_hoy))
            log.info(f"  {etiqueta}: {orig} × {uf_hoy:,.2f} = ${monto_clp:,}")
        else:
            monto_clp = int(valor)
            log.info(f"  {etiqueta}: {orig} = ${monto_clp:,}")

        causa["monto_original"]  = orig
        causa["monto_deuda_clp"] = monto_clp
        extraidos += 1

    log.info("=" * 55)
    log.info("Módulo 3 completado:")
    log.info(f"  Montos extraídos : {extraidos}")
    log.info(f"  Sin PDF          : {sin_pdf}")
    log.info(f"  Fallidos/no enc. : {fallidos}")
    log.info(f"  TOTAL causas     : {len(causas)}")
    log.info("=" * 55)

    return causas


# ─────────────────────────────────────────────────────────────────
# Standalone: dos modos de prueba
# ─────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    import sys

    # Modo --solo: leer PDFs ya descargados en Descargas/ sin correr M1/M2
    if len(sys.argv) > 1 and sys.argv[1] == "--solo":
        causas_test = []
        if not os.path.isdir(DESCARGAS_DIR):
            print(f"Carpeta Descargas no existe: {DESCARGAS_DIR}")
            sys.exit(1)

        for nombre in sorted(os.listdir(DESCARGAS_DIR)):
            if not nombre.endswith(".pdf"):
                continue
            nombre_upper = nombre.upper()
            if "MANDAMIENTO" in nombre_upper:
                tipo_doc  = "mandamiento"
                tipo_proc = "ejecutivo"
            elif "BASES" in nombre_upper:
                tipo_doc  = "bases_remate"
                tipo_proc = "ley_bancos"
            else:
                continue

            # Formato esperado: "C-XXXXX-YYYY_MANDAMIENTO.pdf"
            base = nombre.replace(".pdf", "").split("_")[0]  # "C-32342-2015"
            partes = base.split("-")
            if len(partes) < 3 or partes[0] != "C":
                continue
            rol, año = partes[1], partes[2]

            causas_test.append({
                "rol":               rol,
                "año":               año,
                "tipo_documento":    tipo_doc,
                "tipo_procedimiento": tipo_proc,
                "descargado":        True,
                "ruta_pdf":          os.path.join(DESCARGAS_DIR, nombre),
            })

        print(f"Modo --solo: {len(causas_test)} PDF(s) en {DESCARGAS_DIR}")
        if not causas_test:
            print("No se encontraron PDFs. Ejecuta primero el Módulo 2.")
            sys.exit(0)

        causas_test = extraer_montos(causas_test)
        print(f"\n{'='*55}")
        print("RESULTADO MÓDULO 3 (modo --solo):")
        for c in causas_test:
            if c.get("monto_deuda_clp"):
                print(f"  C-{c['rol']}-{c['año']} | ${c['monto_deuda_clp']:,} | {c['monto_original']}")
            else:
                print(f"  C-{c['rol']}-{c['año']} | SIN MONTO")

    else:
        # Modo pipeline completo M1 → M2 → M3
        from modulo1_parser import parsear_diarios
        from modulo2_ojv import procesar_causas_ojv

        print("Pipeline M1 → M2 → M3")
        causas = parsear_diarios()
        print(f"M1: {len(causas)} causas")

        causas = procesar_causas_ojv(causas)
        desc = sum(1 for c in causas if c.get("descargado"))
        print(f"M2: {desc} documentos descargados")

        causas = extraer_montos(causas)
        print(f"\n{'='*55}")
        print("RESULTADO MÓDULO 3:")
        for c in causas:
            if c.get("monto_deuda_clp"):
                estado = f"${c['monto_deuda_clp']:,}"
            elif c.get("descargado"):
                estado = "SIN MONTO (parseado fallido)"
            else:
                estado = "SIN PDF"
            print(f"  C-{c['rol']}-{c['año']} | {estado} | {c.get('monto_original', '')}")
