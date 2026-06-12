"""
Biblioteca compartida de utilidades de parsing (regex de ROL/tribunal/dirección,
matching de cortes, segmentación). Consumida por v2 (modulo1_v2) y por el
harness test_cbr_docx. Los puntos de entrada v1 (Sonnet/PDF) fueron eliminados
en la Tanda D 2026-06.
"""

import re
import logging
from rapidfuzz import fuzz
import pandas as pd

from config import (
    CAUSAS_XLSX,
    SHEET_REFERENCIA, SHEET_CAUSAS,
)

logging.basicConfig(level=logging.INFO, format="%(asctime)s [M1] %(message)s")
log = logging.getLogger(__name__)

# ────────────────────────────────────────────────────────────────
# Mapeo de ordinales en palabras → número con símbolo °
# ────────────────────────────────────────────────────────────────
ORDINALES = {
    "primer":   "1°",  "primera":   "1°",
    "segundo":  "2°",  "segunda":   "2°",
    "tercer":   "3°",  "tercero":   "3°",  "tercera":  "3°",
    "cuarto":   "4°",  "cuarta":    "4°",
    "quinto":   "5°",  "quinta":    "5°",
    "sexto":    "6°",  "sexta":     "6°",
    "séptimo":  "7°",  "séptima":   "7°",  "septimo":  "7°",  "septima":  "7°",
    "octavo":   "8°",  "octava":    "8°",
    "noveno":   "9°",  "novena":    "9°",
    "décimo":   "10°", "décima":    "10°", "decimo":   "10°", "decima":   "10°",
    "undécimo": "11°", "decimoprimero": "11°",
    "duodécimo": "12°", "decimosegundo": "12°",
    "decimotercero": "13°", "decimocuarto": "14°",
    "decimoquinto":  "15°", "decimosexto":  "16°",
    "decimoséptimo": "17°", "decimoctavo":  "18°",
    "decimonoveno":  "19°",
    "vigésimo":  "20°", "vigesimo":  "20°",
    "vigésima":  "20°", "vigesima":  "20°",
    "trigésimo": "30°",
}

# Compuesto: "vigésimo segundo" → "22°"
DECENAS_COMP = {
    "vigésimo": 20, "vigesimo": 20,
    "trigésimo": 30, "trigesimo": 30,
}
UNIDADES_COMP = {
    "primer": 1, "primero": 1, "primera": 1,
    "segundo": 2, "segunda": 2,
    "tercer": 3, "tercero": 3, "tercera": 3,
    "cuarto": 4, "cuarta": 4,
    "quinto": 5, "quinta": 5,
    "sexto": 6, "sexta": 6,
    "séptimo": 7, "septimo": 7,
    "octavo": 8, "octava": 8,
    "noveno": 9, "novena": 9,
}

# ────────────────────────────────────────────────────────────────
# Utilidades de texto
# ────────────────────────────────────────────────────────────────

def limpiar_texto(texto: str) -> str:
    """Elimina soft-hyphens, normaliza comillas y espacios/saltos de línea."""
    texto = texto.replace("\xad", "")      # soft hyphen (U+00AD)
    texto = texto.replace("\u00ad", "")    # soft hyphen (variante)
    # Normalizar comillas tipográficas → comillas rectas
    texto = texto.replace("\u201c", '"').replace("\u201d", '"')  # " "
    texto = texto.replace("\u2018", "'").replace("\u2019", "'")  # ' '
    texto = texto.replace("\u00ab", '"').replace("\u00bb", '"')  # « »
    texto = re.sub(r"\n+", " ", texto)     # newlines → espacio
    texto = re.sub(r" {2,}", " ", texto)   # múltiples espacios → uno
    # Reconstruir palabras partidas por layout de columnas del PDF
    for palabra_partida, palabra_entera in [
        (r"Juz\s+gado", "Juzgado"),
        (r"JUZGA\s+DO", "JUZGADO"),
        (r"Tri\s+bunal", "Tribunal"),
        (r"BAN\s+CO\b", "BANCO"),
        (r"Con\s+cepci", "Concepci"),
        (r"Valpara\s+[ií]so", "Valparaíso"),
        (r"esta\s+cion", "estacion"),
        (r"Su\s+basta", "Subasta"),
        (r"Se\s+cretaria", "Secretaria"),
        (r"Se\s+cretaría", "Secretaría"),
    ]:
        texto = re.sub(palabra_partida, palabra_entera, texto, flags=re.IGNORECASE)
    return texto.strip()


def normalizar_ordinal(texto: str) -> str:
    """
    Convierte ordinales en palabras a formato numérico con °.
    Ej: "Vigésimo Segundo Juzgado" → "22° Juzgado"
    """
    palabras = texto.split()
    resultado = []
    i = 0
    while i < len(palabras):
        p1 = palabras[i].lower().rstrip(".,;:")
        # Intenta compuesto de dos palabras
        if i + 1 < len(palabras):
            p2 = palabras[i + 1].lower().rstrip(".,;:")
            decena = DECENAS_COMP.get(p1)
            unidad = UNIDADES_COMP.get(p2)
            if decena is not None and unidad is not None:
                resultado.append(f"{decena + unidad}°")
                i += 2
                continue
        # Simple
        if p1 in ORDINALES:
            resultado.append(ORDINALES[p1])
            i += 1
            continue
        resultado.append(palabras[i])
        i += 1
    return " ".join(resultado)


import unicodedata

def _quitar_tildes(texto: str) -> str:
    """Elimina tildes/diacríticos: á→a, é→e, ñ→n, etc."""
    nfkd = unicodedata.normalize("NFKD", texto)
    return "".join(c for c in nfkd if not unicodedata.combining(c))


def _normalizar_para_matching(texto: str) -> str:
    """Normalización agresiva para RapidFuzz."""
    t = texto.lower()
    t = _quitar_tildes(t)
    # Colapsar espacios problemáticos comunes de PDFs
    t = t.replace("val paraiso", "valparaiso")
    t = t.replace("villa rrica", "villarrica")
    t = t.replace("los angeles", "losangeles")
    # Ordinales escritos → numéricos
    t = re.sub(r'\b(primer|primera|1er)\b', '1', t)
    t = re.sub(r'\b(segundo|segunda|2do)\b', '2', t)
    t = re.sub(r'\b(tercer|tercero|tercera|3er)\b', '3', t)
    t = re.sub(r'\b(cuarto|cuarta|4to)\b', '4', t)
    t = re.sub(r'\b(quinto|quinta|5to)\b', '5', t)
    # Abreviaturas estándar
    t = re.sub(r'\bgarantia\b', 'gar', t)
    t = re.sub(r'\bpuerto\b', 'pto', t)
    t = t.replace("jdo", "juzgado")
    # Eliminar símbolo ordinal y puntuación
    t = re.sub(r'[°º\xba\.,-]', ' ', t)
    t = re.sub(r'\bde\b', ' ', t)
    # Solo alfanuméricos y espacios
    t = re.sub(r'[^a-z0-9\s]', ' ', t)
    return re.sub(r'\s+', ' ', t).strip()


# ────────────────────────────────────────────────────────────────
# Carga de datos de referencia
# ────────────────────────────────────────────────────────────────

def cargar_referencia() -> pd.DataFrame:
    """Carga hoja REFERENCIA del Excel (mapeo tribunal → corte)."""
    df = pd.read_excel(CAUSAS_XLSX, sheet_name=SHEET_REFERENCIA)
    df.columns = ["corte", "tribunal"]
    return df


def cargar_historial_roles() -> set:
    """Retorna set de ROLes ya procesados ('XXXXX-YYYY') desde hoja CAUSAS."""
    try:
        df = pd.read_excel(CAUSAS_XLSX, sheet_name=SHEET_CAUSAS)
        if df.empty:
            return set()
        roles = set()
        for _, row in df.iterrows():
            try:
                rol = str(row.get("ROL", "")).strip()
                anio = str(row.get("AÑO", row.get("A\xd1O", ""))).strip()
                if rol and anio:
                    roles.add(f"{rol}-{anio}")
            except Exception:
                continue
        return roles
    except Exception as e:
        log.warning(f"No se pudo leer historial CAUSAS: {e}")
        return set()


# ────────────────────────────────────────────────────────────────
# Limpieza de tribunal (red de seguridad post-LLM)
# ────────────────────────────────────────────────────────────────

def _limpiar_tribunal(nombre_raw: str) -> str:
    """
    Limpia el nombre del tribunal extraído del PDF o de Claude API.
    Se aplica SIEMPRE, como red de seguridad post-LLM.
    """
    if not nombre_raw:
        return nombre_raw

    t = nombre_raw.strip()

    # Regla 4: Unir palabras cortadas por guión de salto de línea
    # "Juzga- do" → "Juzgado", "Corres- pondiente" → "Correspondiente"
    t = re.sub(r'(\w)-\s+(\w)', r'\1\2', t)

    # Regla 5: Eliminar dirección física al final del nombre del tribunal
    # Cortar ante "Nº", "N°", "No.", "Nro" + cualquier cosa
    t = re.sub(r'\s+(Nº|N°|No\.|Nro\.?)\s*\d.*$', '', t, flags=re.IGNORECASE)

    # Cortar ante un número de calle (3+ dígitos al final de la cadena)
    t = re.sub(r'\s+\d{3,}[\s,]?.*$', '', t)

    # Normalizar múltiples espacios
    t = re.sub(r'\s+', ' ', t).strip()

    # Normalizar capitalización: preposiciones en minúscula
    # "Civil De Antofagasta" → "Civil de Antofagasta"
    partes = t.split()
    resultado = []
    for i, p in enumerate(partes):
        if i > 0 and p.lower() in ('de', 'del', 'la', 'las', 'los', 'el'):
            resultado.append(p.lower())
        else:
            resultado.append(p)
    t = ' '.join(resultado)

    return t


_PALABRAS_TRIBUNAL = {
    "juzgado", "civil", "letras", "garantía", "garantia", "mixto", "familia",
    "cobranza", "laboral", "penal", "competencia", "común", "comun", "trabajo",
}

def _extraer_ciudad(nombre_tribunal: str) -> str | None:
    """
    Extrae la ciudad del nombre de un tribunal.
    Ej: '1° Juzgado Civil de Santiago' -> 'santiago'
        'Juzgado de Letras de Rancagua' -> 'rancagua'
        '3° Juzgado de Letras de Iquique' -> 'iquique'  (no 'letras iquique')
    """
    # Buscar la última ocurrencia de "de <Ciudad>" al final
    m = re.search(r'\bde\s+([A-ZÁÉÍÓÚÑ][a-záéíóúñ]+(?:\s+[A-ZÁÉÍÓÚÑ][a-záéíóúñ]+)*)$',
                  nombre_tribunal.strip())
    if m:
        raw = m.group(1)
        # Quitar palabras que son parte del tipo de tribunal, no de la geografía
        palabras = raw.split()
        geo = [p for p in palabras if p.lower() not in _PALABRAS_TRIBUNAL]
        if geo:
            # Normalizar tildes para comparaciones consistentes
            return _quitar_tildes(" ".join(geo).lower())
    return None


def _extraer_ordinal(nombre_tribunal: str) -> str | None:
    """Extrae el número ordinal de un nombre de tribunal (ej: '3° Juzgado...' → '3')."""
    m = re.search(r'(\d+)\s*[°º\xba]', nombre_tribunal)
    if m:
        return m.group(1)
    # Mapeo de ordinales en texto (incluye variantes sin tilde)
    ordinales_texto = {
        'primer': '1', 'primero': '1', 'primera': '1',
        'segundo': '2', 'segunda': '2',
        'tercer': '3', 'tercero': '3', 'tercera': '3',
        'cuarto': '4', 'cuarta': '4',
        'quinto': '5', 'quinta': '5',
        'sexto': '6', 'sexta': '6',
        'séptimo': '7', 'sétimo': '7', 'septimo': '7', 'setimo': '7',
        'octavo': '8', 'octava': '8',
        'noveno': '9', 'novena': '9',
        'décimo': '10', 'decimo': '10',
        'undécimo': '11', 'undecimo': '11',
        'duodécimo': '12', 'duodecimo': '12',
        'decimotercero': '13', 'decimocuarto': '14',
        'decimoquinto': '15', 'decimosexto': '16',
        'decimoséptimo': '17', 'decimoseptimo': '17',
        'decimoctavo': '18', 'decimonoveno': '19',
        'vigésimo': '20', 'vigesimo': '20',
        'trigésimo': '30', 'trigesimo': '30',
    }
    lower = nombre_tribunal.lower()
    # Compuestos: "vigésimo segundo" -> 22
    for dec_txt, dec_val in [('vigésimo', 20), ('vigesimo', 20),
                              ('trigésimo', 30), ('trigesimo', 30)]:
        if dec_txt in lower:
            for uni_txt, uni_val in [('primer', 1), ('primero', 1),
                                      ('segundo', 2), ('tercer', 3), ('tercero', 3),
                                      ('cuarto', 4), ('quinto', 5), ('sexto', 6),
                                      ('séptimo', 7), ('septimo', 7),
                                      ('octavo', 8), ('noveno', 9)]:
                if uni_txt in lower:
                    return str(dec_val + uni_val)
    for texto, num in ordinales_texto.items():
        if texto in lower:
            return num
    return None


# ────────────────────────────────────────────────────────────────
# Matching tribunal → corte
# ────────────────────────────────────────────────────────────────

def buscar_corte(tribunal_raw: str, df_ref: pd.DataFrame,
                 rol: str = "", anio: str = "") -> tuple[str, str]:
    """
    Busca en df_ref la corte que mejor corresponde al tribunal extraído.
    Retorna (corte, tribunal_normalizado).
    Si no encuentra con suficiente confianza, retorna ("DESCONOCIDA", tribunal_raw).
    """
    if not tribunal_raw:
        return "DESCONOCIDA", ""

    etiqueta = f"C-{rol}-{anio}" if rol else ""

    # Normalizar: ordinales, quitar dirección (la dirección suele venir después de coma)
    normalizado = normalizar_ordinal(tribunal_raw)
    # Quitar texto después de la coma (es la dirección del tribunal)
    normalizado = normalizado.split(",")[0].strip()
    # Limpiar ruido habitual
    normalizado = re.sub(r"\s+", " ", normalizado)
    normalizado = normalizado.strip(" .")

    # Post-procesamiento obligatorio: limpiar tribunal (red de seguridad post-LLM)
    normalizado = _limpiar_tribunal(normalizado)

    # Versión normalizada para matching
    norm_match = _normalizar_para_matching(normalizado)

    mejor_score = 0.0
    mejor_corte = "DESCONOCIDA"
    mejor_tribunal = normalizado

    for _, row in df_ref.iterrows():
        trib_ref = str(row["tribunal"])
        ref_match = _normalizar_para_matching(trib_ref)
        # token_set_ratio: ignora orden de palabras y maneja tokens parciales
        score = fuzz.token_set_ratio(norm_match, ref_match)

        if score > mejor_score:
            mejor_score = score
            mejor_corte = str(row["corte"])
            mejor_tribunal = trib_ref

    UMBRAL = 80.0  # RapidFuzz devuelve 0-100; 80 es el punto dulce
    if mejor_score < UMBRAL:
        log.warning(f"  [NO MATCH] {etiqueta} | Tribunal PDF: '{normalizado}' | "
                    f"Mejor Excel: '{mejor_tribunal}' (score: {mejor_score:.1f}%) -> DESCONOCIDA")
        return "DESCONOCIDA", normalizado

    # Validación post-matching: el ordinal del tribunal debe coincidir
    # Evita que "3° de la Serena" matchee a "1° de Letras de la Serena"
    ordinal_pdf = _extraer_ordinal(normalizado)
    ordinal_ref = _extraer_ordinal(mejor_tribunal)
    if ordinal_pdf and ordinal_ref and ordinal_pdf != ordinal_ref:
        log.warning(f"  [ORDINAL MISMATCH] {etiqueta}: PDF dice {ordinal_pdf} pero match es "
                    f"{ordinal_ref} ('{mejor_tribunal}') -- intentando recovery")
        # Recovery: filtrar df_ref solo tribunales con mismo ordinal y re-buscar
        candidatos = []
        for _, row in df_ref.iterrows():
            trib_ref2 = str(row["tribunal"])
            ord_ref2 = _extraer_ordinal(trib_ref2)
            if ord_ref2 == ordinal_pdf:
                candidatos.append((trib_ref2, str(row["corte"])))
        if candidatos:
            mejor_score2 = 0.0
            mejor_corte2 = "DESCONOCIDA"
            mejor_tribunal2 = normalizado
            for trib_ref2, corte_ref2 in candidatos:
                ref_match2 = _normalizar_para_matching(trib_ref2)
                score2 = fuzz.token_set_ratio(norm_match, ref_match2)
                if score2 > mejor_score2:
                    mejor_score2 = score2
                    mejor_corte2 = corte_ref2
                    mejor_tribunal2 = trib_ref2
            if mejor_score2 >= UMBRAL:
                log.info(f"  [ORDINAL RECOVERY] {etiqueta}: re-match '{mejor_tribunal2}' "
                         f"(score: {mejor_score2:.1f}%) | Corte: {mejor_corte2}")
                # Aplicar validación de ciudad al recovery match también
                ciudad_pdf_r = _extraer_ciudad(normalizado)
                ciudad_ref_r = _extraer_ciudad(mejor_tribunal2)
                if ciudad_pdf_r and ciudad_ref_r and ciudad_pdf_r != ciudad_ref_r:
                    score_pen = mejor_score2 * 0.7
                    if score_pen < UMBRAL:
                        log.warning(f"  [CITY MISMATCH post-recovery] {etiqueta}: "
                                    f"'{ciudad_pdf_r}' vs '{ciudad_ref_r}' -- rechazando")
                        return "DESCONOCIDA", normalizado
                    mejor_score2 = score_pen
                return mejor_corte2, mejor_tribunal2
            else:
                log.warning(f"  [ORDINAL RECOVERY FAIL] {etiqueta}: mejor '{mejor_tribunal2}' "
                            f"(score: {mejor_score2:.1f}%) < {UMBRAL}")
        return "DESCONOCIDA", normalizado

    # Validación post-matching: la ciudad del tribunal debe coincidir
    # Evita que "Rancagua" matchee a "Buin" o "Linares" a "Valparaíso"
    ciudad_pdf = _extraer_ciudad(normalizado)
    ciudad_ref = _extraer_ciudad(mejor_tribunal)
    if ciudad_pdf and ciudad_ref and ciudad_pdf != ciudad_ref:
        score_penalizado = mejor_score * 0.7
        if score_penalizado < UMBRAL:
            log.warning(f"  [CITY MISMATCH] {etiqueta}: PDF '{ciudad_pdf}' vs ref '{ciudad_ref}' | "
                        f"score {mejor_score:.1f}% * 0.7 = {score_penalizado:.1f}% < {UMBRAL} -- intentando recovery")
            # Recovery: filtrar df_ref por tribunales que contengan la ciudad del PDF
            candidatos_city = []
            for _, row in df_ref.iterrows():
                trib_ref_c = str(row["tribunal"])
                ciudad_ref_c = _extraer_ciudad(trib_ref_c)
                if ciudad_ref_c and ciudad_ref_c == ciudad_pdf:
                    candidatos_city.append((trib_ref_c, str(row["corte"])))
            if candidatos_city:
                mejor_score_c = 0.0
                mejor_corte_c = "DESCONOCIDA"
                mejor_tribunal_c = normalizado
                for trib_ref_c, corte_ref_c in candidatos_city:
                    ref_match_c = _normalizar_para_matching(trib_ref_c)
                    score_c = fuzz.token_set_ratio(norm_match, ref_match_c)
                    if score_c > mejor_score_c:
                        mejor_score_c = score_c
                        mejor_corte_c = corte_ref_c
                        mejor_tribunal_c = trib_ref_c
                if mejor_score_c >= UMBRAL:
                    # Validar ordinal en el recovery match
                    ord_pdf_c = _extraer_ordinal(normalizado)
                    ord_ref_c = _extraer_ordinal(mejor_tribunal_c)
                    if ord_pdf_c and ord_ref_c and ord_pdf_c != ord_ref_c:
                        log.warning(f"  [CITY RECOVERY ORDINAL FAIL] {etiqueta}: "
                                    f"ordinal {ord_pdf_c} vs {ord_ref_c} -- rechazando")
                        return "DESCONOCIDA", normalizado
                    log.info(f"  [CITY RECOVERY] {etiqueta}: re-match '{mejor_tribunal_c}' "
                             f"(score: {mejor_score_c:.1f}%) | Corte: {mejor_corte_c}")
                    return mejor_corte_c, mejor_tribunal_c
                else:
                    log.warning(f"  [CITY RECOVERY FAIL] {etiqueta}: mejor '{mejor_tribunal_c}' "
                                f"(score: {mejor_score_c:.1f}%) < {UMBRAL}")
            return "DESCONOCIDA", normalizado
        else:
            log.info(f"  [CITY MISMATCH leve] {etiqueta}: PDF '{ciudad_pdf}' vs ref '{ciudad_ref}' | "
                     f"score penalizado {score_penalizado:.1f}% >= {UMBRAL} -- aceptando")
            mejor_score = score_penalizado

    log.info(f"  [MATCH] {etiqueta} | Tribunal: '{normalizado}' -> "
             f"'{mejor_tribunal}' (score: {mejor_score:.1f}%) | Corte: {mejor_corte}")
    return mejor_corte, mejor_tribunal


# ────────────────────────────────────────────────────────────────
# Extracción de campos desde un bloque de texto
# ────────────────────────────────────────────────────────────────

# Patrón ROL judicial: C-XXXXX-YYYY (el número puede tener puntos)
_RE_ROL = re.compile(
    r"[Rr]ol(?:\s*(?:N[°º]|Nro\.?))?\s*C\s*[-–—]?\s*(\d[\d.\s]*\d|\d)\s*[-–—]\s*(\d{4})",
    re.IGNORECASE,
)
# También el formato "EXTRACTO REMATE ROL C-93-2024"
_RE_ROL_HEADER = re.compile(
    r"(?:EXTRACTO\s+)?REMATE\s+ROL\s+C\s*[-–—]?\s*(\d[\d.\s]*\d|\d)\s*[-–—]\s*(\d{4})",
    re.IGNORECASE,
)


def extraer_rol(bloque: str) -> tuple[str, str] | None:
    """Retorna (rol, año) o None si no hay ROL válido."""
    for pat in (_RE_ROL_HEADER, _RE_ROL):
        m = pat.search(bloque)
        if m:
            rol = m.group(1).replace(".", "").replace(" ", "")  # quitar puntos y espacios
            anio = m.group(2)
            return rol, anio
    return None


def extraer_tribunal_texto(bloque: str) -> str:
    """
    Extrae el nombre del tribunal del encabezado del bloque.
    Aplica múltiples estrategias en orden de confianza.
    """
    b = bloque.strip()

    # ── Estrategia 1: "EXTRACTO REMATE ROL C-XXX Ante el Tribunal, ..."
    m = re.match(
        r"EXTRACTO\s+REMATE\s+ROL\s+C-[\d.]+-\d{4}\s+Ante\s+el?\s+(.+?)(?:,|ubicado|se\s+remat)",
        b, re.IGNORECASE
    )
    if m:
        return m.group(1).strip()

    # ── Estrategia 2: "EXTRACTO REMATE: Tribunal, ..."
    m = re.match(
        r"EXTRACTO\s+REMATE[:\s.]+(?:Ante\s+el?\s+)?(.+?)(?:,|\s+ubicado|\s+rematará|\s+rematara|\s+se\s+remat)",
        b, re.IGNORECASE
    )
    if m:
        candidato = m.group(1).strip()
        if _es_tribunal(candidato):
            return candidato

    # ── Estrategia 3: "Remate: Tribunal, ..." o "Remate: Tribunal Rol ..."
    m = re.match(
        r"Remate:\s+(.+?)(?:,|\s+Rol\s+N[°º]|\s+rol\s+C-|\s+se\s+remat)",
        b, re.IGNORECASE
    )
    if m:
        candidato = m.group(1).strip()
        if _es_tribunal(candidato):
            return candidato

    # ── Estrategia 4: "REMATE. Ante [el] Tribunal, ..."
    m = re.match(r"REMATE[. ]+Ante\s+(?:el\s+|la\s+)?(.+?)(?:,|\s+ubicado)", b, re.IGNORECASE)
    if m:
        return m.group(1).strip()

    # ── Estrategia 5: "REMATE JUDICIAL. Tribunal" o "REMATE JUDICIAL TRIBUNAL"
    # Formato corto: "REMATE JUDICIAL. TERCER Juzgado Civil Concepción, ..."
    m = re.match(
        r"REMATE\s+JUDICIAL[.\s]+(?!Causa\s|En\s|\.{2})([A-ZÁÉÍÓÚÑ][\w\s.°º]+?(?:[Jj]uzgado|JUZGADO)[^,\n]{0,60}?)(?:,|\s+[Cc]ausa\s+[Rr]ol|\s+[Ee]n\s+autos|\s+[Oo]rden[oó]|\s+[Ee]n\s+causa)",
        b, re.IGNORECASE
    )
    if m:
        candidato = normalizar_ordinal(m.group(1).strip())
        if _es_tribunal(candidato):
            return candidato

    # ── Estrategia 6: "REMATE JUDICIAL Causa Rol C-XXX «caratulados» Tribunal ordenó"
    # Formato: "REMATE JUDICIAL Causa Rol C-3676-2020, «LAURA/CANALES», Tercer Juzgado Letras Iquique, Ordenó..."
    # (tras limpiar_texto las comillas son comillas rectas normales)
    m = re.search(
        r'"[^"]{3,}?"\s*,?\s*([A-ZÁÉÍÓÚÑ][^"\n,]{5,80}?[Jj]uzgado[^"\n,]{0,50}?)\s*(?:,?\s*[Oo]rden[oó]|\s+\d+\s+de|\s+del\s+d)',
        b, re.IGNORECASE
    )
    if m:
        return normalizar_ordinal(m.group(1).strip())

    # ── Estrategia 7: "REMATE JUDICIAL [ALL CAPS TRIBUNAL] En autos/causa"
    # Formato: "REMATE JUDICIAL PRIMER JUZGADO LA SERENA En autos ejecutivos..."
    m = re.match(
        r"REMATE\s+JUDICIAL\s+([A-Z][A-Z\s]+?(?:JUZGADO|TRIBUNAL)[A-Z\s.°º]*?)\s+(?:[Ee]n\s+autos|[Cc]ausa\s+[Rr]ol|[Ee]n\s+causa)",
        b
    )
    if m:
        return normalizar_ordinal(m.group(1).strip())

    # ── Estrategia 8: "REMATE. 3° Juzgado..." / "REMATE. TERCER JUZGADO..."
    # Incluye "REMATE SEGUNDO. Juzgado..." y "REMATE. 3° Juzgado de Letras..."
    m = re.match(
        r"REMATE[. ]+(\d+[°º\xba]?\s*[Jj]uzgado[^,\n]{0,60}?|[A-ZÁÉÍÓÚÑ][\w\s.°º\xba]*?[Jj]uzgado[^,\n]{0,60}?)"
        r"(?:,|\s+en\s+causa|\s+juicio|\s+[Rr]ol\s+N|\s+se\s+remat|\s+ubicado|\s+[Cc]alle|\s+\d+\s+de|[.]\s)",
        b, re.IGNORECASE
    )
    if m:
        candidato = normalizar_ordinal(m.group(1).strip(" ."))
        if _es_tribunal(candidato):
            return candidato

    # ── Estrategia 9: En el cuerpo "seguida ante el Tribunal"
    m = re.search(r"seguida?\s+ante\s+el?\s+(.+?)(?:,|\s+ubicado)", b, re.IGNORECASE)
    if m:
        return m.group(1).strip()

    # ── Estrategia 10: "en Secretaría/sala del Tribunal" / "en dependencias del Tribunal"
    m = re.search(
        r"en\s+(?:[Ss]ecretar[ií]a\s+del\s+|[Ss]ala\s+del\s+|dependencias\s+del\s+)([^,.\n]{5,80}?[Jj]uzgado[^,.\n]{0,50}?)(?:,|\.|\s+como\s+|\s+el\s+\d|\s+se\s+remat)",
        b, re.IGNORECASE
    )
    if m:
        return normalizar_ordinal(m.group(1).strip())

    # ── Estrategia 11: Buscar "Juzgado" mencionado en el texto del bloque (fallback)
    m = re.search(
        r"(?:del|ante\s+el?|el)\s+((?:\d+[°º]\s+|[A-ZÁÉÍÓÚÑ][a-záéíóúñ]+\s+)?[Jj]uzgado\s+(?:[Cc]ivil|[Dd]e\s+[Ll]etras)[^,\n.]{0,50})",
        b
    )
    if m:
        return normalizar_ordinal(m.group(1).strip())

    return ""


def _es_tribunal(texto: str) -> bool:
    """Verifica que el texto extraído parezca un tribunal (contiene 'juzgado' o 'tribunal')."""
    t = texto.lower()
    return any(k in t for k in ("juzgado", "tribunal", "letras", "civil"))


def extraer_demandante(bloque: str) -> str:
    """
    Extrae el nombre del demandante (la parte antes de 'con' en los caratulados).
    Nota: el bloque ya pasó por limpiar_texto(), comillas normalizadas a "".
    """
    # Patrón: caratulado[s] "DEMANDANTE con DEMANDADO"
    m = re.search(
        r'caratulad[ao]s?\s*"([^"]+?)\s+con\s+[A-ZÁÉÍÓÚÑ]',
        bloque, re.IGNORECASE
    )
    if m:
        return m.group(1).strip().strip("\"' ")

    # Patrón: "en causa DEMANDANTE con DEMANDADO" (sin comillas)
    # NOTA: NO usar re.IGNORECASE aquí — el [A-ZÁÉÍÓÚÑ] final debe ser mayúscula
    # para distinguir "con DEMANDADO" de "con a lo menos", "con fecha", etc.
    m = re.search(
        r'[Cc]ausa\s+(?:[Rr]ol\s+(?:[Nn][°º\xba])?\s*C-[\d.]+-\d{4}\s+)?"?([A-ZÁÉÍÓÚÑ][^"\n]{3,80}?)\s+[Cc][Oo][Nn]\s+[A-ZÁÉÍÓÚÑ]',
        bloque
    )
    if m:
        candidato = m.group(1).strip().strip("\"'")
        if len(candidato) > 3 and not re.fullmatch(r"C-[\d.]+-\d{4}", candidato):
            return candidato

    # Patrón: "Causa DEMANDANTE/DEMANDADO" (slash en vez de con)
    m = re.search(
        r'[Cc]ausa\s+([A-ZÁÉÍÓÚÑ][^/\n]{3,50}?)\s*/\s*[A-ZÁÉÍÓÚÑ]',
        bloque
    )
    if m:
        return m.group(1).strip()

    # Patrón: "BancoXxx/Apellido" directamente (sin "causa")
    m = re.search(
        r'\b([A-ZÁÉÍÓÚÑ][A-Za-záéíóúñüÜ\s.]+?(?:banco|banca|corp|vida|seguro)[A-Za-záéíóúñ\s.]*?)\s*/\s*\w',
        bloque, re.IGNORECASE
    )
    if m:
        return m.group(1).strip()

    return ""


# Patrones de dirección
# Terminadores comunes en escrituras chilenas:
#   "inscrito", "el dominio", "inmueble", "con acceso"  → ya estaban
#   "que corresponde"  → corta tras "calle X número N, que corresponde al sitio..."
#   "resultante de"    → corta tras "calle X, resultante de la subdivisión..."
_RE_DIRECCION = [
    # "ubicado en DIRECCIÓN"
    re.compile(r"ubicad[ao]\s+en\s+(.{10,150}?)(?:\.|,\s*(?:inscrito|el\s+dominio|inmueble|con\s+acceso|que\s+corresponde|resultante\s+de)|$)",
               re.IGNORECASE),
    # "acceso [principal/común] por DIRECCIÓN"
    re.compile(r"acceso\s+(?:principal\s+|com[uú]n\s+)?por\s+(.{10,120}?)(?:,|\.|$)",
               re.IGNORECASE),
    # "propiedad ubicada en DIRECCIÓN"
    re.compile(r"propiedad\s+(?:ubicada?\s+en\s+)(.{10,150}?)(?:\.|,\s*(?:inscrito|el\s+dominio|que\s+corresponde|resultante\s+de)|$)",
               re.IGNORECASE),
    # "inmueble [consistente en] [el ...] ubicado en DIRECCIÓN"
    re.compile(r"inmueble\s+(?:consistente[^,]+,\s+)?(?:[^,]+,\s+)?(?:con\s+acceso\s+(?:por|en)\s+)?(.{5,120}?)(?:,\s*(?:inscrito|el\s+dominio|ciudad)|$)",
               re.IGNORECASE),
]

_RE_COMUNA = re.compile(
    r"(?:ciudad\s+y\s+)?(?:ciudad,\s+)?comuna\s+(?:y\s+\w+\s+)?(?:de\s+la\s+)?(?:de\s+)?([A-ZÁÉÍÓÚÑ][A-Za-záéíóúñ\s]+?)(?:\s*,|\s*\.|$|\s+Regi[oó]n|\s+Provincia)",
    re.IGNORECASE
)

_RE_CALLE_DIRECTA = re.compile(
    r"\b(?:calle|pasaje|avenida|av\.|psje\.)\s+[A-ZÁÉÍÓÚÑ].{5,100}?N[°º]?\s*\d+",
    re.IGNORECASE
)


def extraer_tipo_propiedad(bloque: str) -> str:
    """
    Detecta el tipo de propiedad mencionado en el aviso.

    Reglas (en orden de prioridad):
      "condominio"              → "Ambos"   (unidades pueden ser casa o depto)
      "departamento" / "dpto"   → "Departamentos"
      cualquier otro caso       → "Casas"   (default seguro)
    """
    b = bloque.lower()
    if re.search(r"\bcondominio\b", b):
        return "Ambos"
    if re.search(r"\bdepartamento\b|\bdpto\.?\b", b):
        return "Departamentos"
    return "Casas"


# Palabras clave que indican que el "ubicado en" es la dirección del TRIBUNAL, no del inmueble
_TRIBUNAL_ADDR_CONTEXT = re.compile(
    r"(?:juzgado|tribunal|letras|civil|secretar[ií]a|oficial)\s+(?:\w+\s+){0,5}ubicad",
    re.IGNORECASE
)
# Direcciones conocidas de tribunales (Santiago principalmente)
_TRIBUNAL_ADDR_KNOWN = re.compile(
    r"(?:hu[eé]rfanos|rengifo|barros\s+arana|san\s+mart[ií]n\s+N[°º]?\s*2984|santiago\s+trigo|bilbao\s+777|ciro\s+arredondo|errazuriz\s+s/n)",
    re.IGNORECASE
)


def extraer_direccion_comuna(bloque: str) -> tuple[str, str]:
    """
    Retorna (dirección, comuna) del inmueble.
    Intenta múltiples patrones. Si no encuentra, retorna cadenas vacías.
    Excluye direcciones de tribunales.
    """
    direccion = ""
    comuna = ""

    # Intentar encontrar la fecha de la subasta para buscar la dirección del inmueble
    # DESPUÉS de ese punto (ya que la dirección del tribunal suele aparecer antes)
    m_fecha = re.search(
        r"(?:se\s+remat[aá]r[aá]?|rematar[aá]\s+el\s+d[ií]a|el\s+d[ií]a\s+\d+\s+de\s+\w+\s+de\s+20\d\d)",
        bloque, re.IGNORECASE
    )
    # Buscar desde la fecha si se encontró, si no desde la mitad del bloque
    offset_busqueda = m_fecha.start() if m_fecha else max(0, len(bloque) // 4)

    # Buscar dirección en el texto a partir del offset
    for pat in _RE_DIRECCION:
        for m in pat.finditer(bloque):
            if m.start() < offset_busqueda:
                # Antes del offset, verificar que no sea dirección de tribunal
                contexto_previo = bloque[max(0, m.start()-100):m.start()]
                if _TRIBUNAL_ADDR_CONTEXT.search(contexto_previo):
                    continue
            candidato = m.group(1).strip().strip(",. ")
            # Filtrar candidatos muy cortos
            if len(candidato) < 8:
                continue
            # Filtrar direcciones conocidas de tribunales
            if _TRIBUNAL_ADDR_KNOWN.search(candidato[:60]):
                continue
            # Filtrar si el candidato empieza con términos del tribunal
            if re.match(r"^(?:piso|oficina|sala|secretar)", candidato, re.IGNORECASE):
                continue
            direccion = candidato
            break
        if direccion:
            break

    # Si no encontramos dirección con _RE_DIRECCION, intentar con patrón de calle directa
    if not direccion:
        for m in _RE_CALLE_DIRECTA.finditer(bloque):
            candidato = m.group(0).strip()
            if not _TRIBUNAL_ADDR_KNOWN.search(candidato[:60]):
                direccion = candidato
                break

    # Extraer comuna — buscar todas las ocurrencias y elegir la más relevante
    comunas_encontradas = list(_RE_COMUNA.finditer(bloque))
    for mc in comunas_encontradas:
        candidato_comuna = mc.group(1).strip().strip(",. ")
        # Limpiar residuos
        for ruido in ["inscrito", "el dominio", "provincia", "region", "región"]:
            if ruido in candidato_comuna.lower():
                candidato_comuna = candidato_comuna[:candidato_comuna.lower().find(ruido)].strip(", ")
        if len(candidato_comuna) >= 3:
            # Preferir la primera comuna que aparece después del offset
            if mc.start() >= offset_busqueda or not comuna:
                comuna = candidato_comuna
                if mc.start() >= offset_busqueda:
                    break

    # Limpiar dirección (quitar "commune" al final si quedó pegada)
    if direccion and "comuna" in direccion.lower():
        idx = direccion.lower().find("comuna")
        direccion = direccion[:idx].strip(", ")

    # Truncar dirección muy larga
    if len(direccion) > 150:
        # Cortar en la primera coma si es muy largo
        partes = direccion.split(",")
        direccion = partes[0].strip() if partes else direccion[:150]

    return direccion, comuna


# ────────────────────────────────────────────────────────────────
# Separación de bloques en el texto del PDF
# ────────────────────────────────────────────────────────────────

_RE_BLOQUE_SEP = re.compile(
    r"(?:EXTRACTO\s+REMATE|REMATE\s+JUDICIAL|REMATE[. ]|Remate:)",
)


def separar_bloques(texto: str) -> list[str]:
    """
    Divide el texto completo del PDF en bloques individuales de remate.
    Cada bloque empieza en un marcador de remate.
    """
    matches = list(_RE_BLOQUE_SEP.finditer(texto))
    bloques = []
    for i, m in enumerate(matches):
        inicio = m.start()
        fin = matches[i + 1].start() if i + 1 < len(matches) else len(texto)
        bloques.append(texto[inicio:fin])
    return bloques


# ────────────────────────────────────────────────────────────────
# Regex de segmentación del DOCX (usados por test_cbr_docx.py)
# ────────────────────────────────────────────────────────────────

_MESES_ES = (
    r'(?:ENERO|FEBRERO|MARZO|ABRIL|MAYO|JUNIO|JULIO|AGOSTO|'
    r'SEPTIEMBRE|OCTUBRE|NOVIEMBRE|DICIEMBRE)'
)
_RE_FECHA_ENCABEZADO = re.compile(rf'^\d{{1,2}}\s+(?:DE\s+)?{_MESES_ES}$')
_RE_TITULO_SEMANA = re.compile(r'REMATES\s+SEMANA', re.IGNORECASE)


# ────────────────────────────────────────────────────────────────
# Standalone deshabilitado (Tanda D 2026-06)
# ────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    print("Modo standalone v1 deshabilitado (Tanda D). Usa los flags "
          "del módulo o el pipeline completo: python main.py --docx ...")
    raise SystemExit(0)