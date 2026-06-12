"""
Filtro de Antigüedad CBR (Audit 11)
Heurística para detectar año de inscripción del dominio en el Conservador
de Bienes Raíces y decidir si excluir, mantener o mandar a revisión manual.

Regla: dominio inscrito >= 2020 → EXCLUIR (ahorro API)
       dominio inscrito < 2020  → MANTENER
       año no detectado/ambiguo → REVISAR (va a PENDIENTE_REVISION_MANUAL)
"""

import re

_CBR_ANIO_CORTE = 2020
_CBR_ANIO_MIN   = 1900
_CBR_ANIO_MAX   = 2027
_CBR_VENTANA    = 150

_CBR_ANCLAS = re.compile(
    r"registro\s+(?:de\s+)?propiedad"
    r"|conservador"
    r"|bienes\s+ra[ií]ces"
    r"|c\.?\s*b\.?\s*r\.?"
    r"|inscri(?:t[oa]|pci[oó]n)"
    r"|fojas?", re.IGNORECASE)
_CBR_ANIO_TAG = re.compile(r"a[ñn]o(?:\s+de)?\s+(\d{4})", re.IGNORECASE)
_CBR_VIGENTE  = re.compile(r"actual|vigente", re.IGNORECASE)


def evaluar_antiguedad_cbr(bloque_texto: str) -> dict:
    """
    Evalúa la antigüedad del dominio según el año de inscripción en el CBR.

    Retorna un dict con:
      - decision: "EXCLUIR", "MANTENER", o "REVISAR"
      - cbr_anio: int o None
      - cbr_flag_revision: bool
      - cbr_motivo: str (glosa si va a revisión, vacío en otro caso)
    """
    def _revisar(motivo):
        return {"decision": "REVISAR", "cbr_anio": None,
                "cbr_flag_revision": True, "cbr_motivo": motivo}

    texto = re.sub(r"\s+", " ", (bloque_texto or "").lower())
    anclas = [m.start() for m in _CBR_ANCLAS.finditer(texto)]
    if not anclas:
        return _revisar("Año CBR no detectado")

    candidatos = []
    for m in _CBR_ANIO_TAG.finditer(texto):
        a = int(m.group(1))
        if not (_CBR_ANIO_MIN <= a <= _CBR_ANIO_MAX):
            continue
        pos = m.start()
        if any(abs(pos - ap) <= _CBR_VENTANA for ap in anclas):
            candidatos.append((a, pos))

    if not candidatos:
        return _revisar("Año CBR no detectado")

    valores = {a for a, _ in candidatos}
    if len(valores) == 1:
        anio = next(iter(valores))
    else:
        marca = _CBR_VIGENTE.search(texto)
        if marca:
            anio = min(candidatos, key=lambda c: abs(c[1] - marca.start()))[0]
        else:
            return _revisar("Año CBR ambiguo")

    if anio >= _CBR_ANIO_CORTE:
        return {"decision": "EXCLUIR", "cbr_anio": anio,
                "cbr_flag_revision": False, "cbr_motivo": ""}
    return {"decision": "MANTENER", "cbr_anio": anio,
            "cbr_flag_revision": False, "cbr_motivo": ""}