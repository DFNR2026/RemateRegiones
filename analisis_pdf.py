"""Análisis de PDFs de actas de remate - extraído de filtrador_saldos.py (Frente B)."""

import io
import os
import re

import fitz
import pytesseract
from PIL import Image

_CARGO_REGEX = re.compile(r'cargo\s+(?:a\s+(?:su|sus|los?)|al)\s+cr[eé]ditos?', re.IGNORECASE)

# Regex para montos CLP: $XX.XXX.XXX o $XX.XXX.XXX.- (punto millar chileno)
_REGEX_MONTO_CLP = re.compile(r'\$\s*([\d]+(?:\.[\d]{3})*(?:,[\d]+)?)\s*(?:\.\-|\.-)?')


def _extraer_texto_pdf(ruta):
    """Extrae texto de un PDF con PyMuPDF. Retorna string vacio si falla."""
    from filtrador_saldos import log

    try:
        doc = fitz.open(ruta)
        texto = ""
        for pagina in doc:
            texto += pagina.get_text()
        doc.close()
        return texto
    except Exception as e:
        log.warning("Error leyendo PDF %s: %s", ruta, e)
        return ""


def _ocr_pdf(doc):
    """Extrae texto de PDF escaneado usando OCR (Tesseract).

    Args:
        doc: fitz.Document abierto
    Returns:
        str con texto extraído, o "" si falla
    """
    from filtrador_saldos import log

    try:
        # Configurar ruta de Tesseract en Windows
        tesseract_paths = [
            r"C:\Program Files\Tesseract-OCR\tesseract.exe",
            r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
            r"C:\Users\ndieg\AppData\Local\Programs\Tesseract-OCR\tesseract.exe",
        ]
        for path in tesseract_paths:
            if os.path.exists(path):
                pytesseract.pytesseract.tesseract_cmd = path
                # Fix OCR actas [2026-05-24]: Tesseract no encontraba spa.traineddata
                # porque TESSDATA_PREFIX no estaba seteado (buscaba en ruta relativa
                # D:\Remates\.tessdata). Apuntar a la carpeta tessdata de la instalacion.
                tessdata_dir = os.path.join(os.path.dirname(path), "tessdata")
                if os.path.isdir(tessdata_dir):
                    os.environ["TESSDATA_PREFIX"] = tessdata_dir
                else:
                    log.warning("  tessdata no encontrado en %s -- OCR puede fallar", tessdata_dir)
                break
        else:
            # Ninguna ruta de tesseract existe: avisar fuerte (antes fallaba en silencio)
            log.warning("  Tesseract.exe NO encontrado en rutas conocidas -- OCR no disponible")

        texto_total = ""
        total_paginas = len(doc)
        paginas_fallidas = 0
        for page_num in range(total_paginas):
            try:
                page = doc[page_num]
                pix = page.get_pixmap(dpi=150)  # 150 DPI: más rápido, suficiente calidad
                img_bytes = pix.tobytes("png")
                img = Image.open(io.BytesIO(img_bytes))
                texto_pagina = pytesseract.image_to_string(img, lang='spa', timeout=30)
                texto_total += texto_pagina + "\n"
            except Exception as e:
                paginas_fallidas += 1
                log.warning("    OCR pagina %d fallo: %s", page_num, e)
                continue

        # Fix OCR actas [2026-05-24]: si TODAS las paginas fallan, OCR esta roto
        # (antes pasaba inadvertido con warnings discretos por pagina).
        if total_paginas > 0 and paginas_fallidas == total_paginas:
            log.error("  OCR fallo en TODAS las paginas (%d). Revisar TESSDATA_PREFIX/idioma.", total_paginas)

        return texto_total

    except ImportError:
        log.warning("  pytesseract no instalado. Instalar: pip install pytesseract Pillow")
        return ""
    except Exception as e:
        log.warning("  Error OCR: %s", e)
        return ""


def _analizar_pdf_acta(filepath):
    """Analiza un PDF de acta de remate. Usa OCR si el PDF es imagen escaneada.

    Busca:
    1. "cargo a su crédito" -> banco se adjudicó, sin excedente posible
    2. Monto de adjudicación en CLP (posicional: cerca de "adjudica")

    Returns:
        dict con:
            'cargo_al_credito': bool
            'monto_adjudicacion': int|None
            'texto_monto': str|None (texto original del monto, para log)
    """
    from filtrador_saldos import log

    resultado = {
        "cargo_al_credito": False,
        "monto_adjudicacion": None,
        "texto_monto": None,
    }
    try:
        doc = fitz.open(filepath)

        # Paso 1: Intentar extracción de texto nativo
        texto = ""
        for pagina in doc:
            texto += pagina.get_text()

        # Paso 2: Detectar si es PDF escaneado por CONTENIDO, no por longitud
        # Un acta de remate con texto nativo SIEMPRE contiene "$" y "adjudica" o "remate"
        texto_lower = texto.lower()
        tiene_contenido_acta = (
            ('$' in texto and any(c.isdigit() for c in texto)) or
            ('adjudica' in texto_lower) or
            ('suma de' in texto_lower and any(c.isdigit() for c in texto))
        )

        if not tiene_contenido_acta:
            log.info("    PDF sin contenido de acta en texto nativo. Intentando OCR...")
            texto_ocr = _ocr_pdf(doc)
            if texto_ocr and len(texto_ocr.strip()) > 50:
                log.info("    OCR exitoso: %d chars extraidos", len(texto_ocr))
                texto = texto_ocr
            else:
                log.info("    OCR fallo o sin texto util")
                doc.close()
                return resultado

        doc.close()
    except Exception as e:
        log.warning("  Error leyendo PDF acta: %s", e)
        return resultado

    if not texto.strip():
        return resultado

    texto_lower = texto.lower()

    # === 1. Detectar "cargo al credito" (regex flexible) ===
    if _CARGO_REGEX.search(texto):
        resultado["cargo_al_credito"] = True

    # === 2. Extraer monto de adjudicacion (posicional) ===
    # Buscar todos los montos >= 1 millon en el texto
    all_montos = []
    for match in _REGEX_MONTO_CLP.finditer(texto):
        monto_str = match.group(1).replace(".", "").replace(",", "")
        try:
            monto_int = int(monto_str)
            if monto_int >= 1_000_000:
                all_montos.append({
                    "valor": monto_int,
                    "texto": match.group(0),
                    "pos": match.start(),
                })
        except ValueError:
            continue

    if all_montos:
        # Buscar posicion de "adjudica" en el texto (ultima ocurrencia)
        adjudica_pos = -1
        for keyword in ["adjudica", "adjudicó", "adjudicacion"]:
            pos = texto_lower.rfind(keyword)
            if pos > adjudica_pos:
                adjudica_pos = pos

        if adjudica_pos >= 0:
            # Tomar primer monto DESPUES de la ultima "adjudica"
            montos_post = [m for m in all_montos if m["pos"] > adjudica_pos]
            if montos_post:
                mejor = montos_post[0]
            else:
                # Fallback: monto mas grande del documento
                mejor = max(all_montos, key=lambda m: m["valor"])
        else:
            # Sin "adjudica" -> monto mas grande
            mejor = max(all_montos, key=lambda m: m["valor"])

        resultado["monto_adjudicacion"] = mejor["valor"]
        resultado["texto_monto"] = mejor["texto"]

    # Tope de sanidad: descartar montos irrazonables (OCR puede juntar líneas)
    MONTO_MAXIMO_RAZONABLE = 2_000_000_000  # $2 mil millones CLP
    if resultado["monto_adjudicacion"] and resultado["monto_adjudicacion"] > MONTO_MAXIMO_RAZONABLE:
        log.warning("    ALERTA: Monto $%s supera tope razonable. Descartando.",
                     f"{resultado['monto_adjudicacion']:,}")
        resultado["monto_adjudicacion"] = None
        resultado["texto_monto"] = None

    return resultado