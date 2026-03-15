"""
Módulo 2: Consulta OJV + Descarga de documentos

Adapta ojv_remates.py (v10.0) para recibir la lista del Módulo 1
en vez de leer el Excel directamente.

Input:  lista de dicts de modulo1_parser (rol, año, corte, tribunal, ...)
Output: misma lista enriquecida con:
        - tipo_procedimiento : "ejecutivo" | "ley_bancos" | "desposeimiento" | ""
        - tipo_documento     : "mandamiento" | "bases_remate" | ""
        - descargado         : True | False
        - ruta_pdf           : ruta al PDF descargado o ""
"""

import os
import re
import sys
import time
import logging

# Forzar UTF-8 en stdout/stderr para que los print() de ojv_remates.py
# (que usan ✓ ✗ → ⚠ etc.) no fallen en terminales Windows con cp1252.
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

from playwright.sync_api import sync_playwright

# Importar helpers de ojv_remates (ya probados en v10.0)
from ojv_remates import (
    cerrar_popups,
    cerrar_modal_aviso,
    seleccionar_por_texto,
    navegar_a_consulta,
    limpiar_formulario,
    buscar_causa,
    abrir_detalle,
    seleccionar_cuaderno,
    filas_del_modal,
    descargar_pdf_de_fila,
    buscar_mandamiento,
    buscar_bases_remate,
)

from config import DESCARGAS_DIR, CAUSAS_IGNORADAS

logging.basicConfig(level=logging.INFO, format="%(asctime)s [M2] %(message)s")
log = logging.getLogger(__name__)


# ─────────────────────────────────────────────────────────────────
# Parche: redirigir CARPETA_DESCARGAS de ojv_remates al valor de config
# ─────────────────────────────────────────────────────────────────
import ojv_remates as _ojv
_ojv.CARPETA_DESCARGAS = DESCARGAS_DIR


# ─────────────────────────────────────────────────────────────────
# Extracción de litigantes (DTE + DDO) desde pestaña #litigantesCiv
# ─────────────────────────────────────────────────────────────────

def _extraer_litigantes_ojv(page, etiqueta: str) -> dict:
    """
    Hace click en la pestaña Litigantes del modal, lee SOLO la tabla
    #litigantesCiv y extrae DTE (demandante) y DDO (demandado).

    Columnas reales de la OJV:
      celdas[0] = Participante ("DTE.", "DDO.")
      celdas[1] = RUT
      celdas[2] = Tipo persona ("NATURAL", "JURIDICA")
      celdas[3] = Nombre completo

    Returns: {'nombre_dte': str|None, 'nombre_ddo': str|None}
    Siempre vuelve a la pestaña Historia al terminar.
    """
    resultado = {'nombre_dte': None, 'nombre_ddo': None}

    try:
        # 1. Click en pestaña Litigantes
        tab_link = page.locator('a[href="#litigantesCiv"]')
        if tab_link.count() == 0:
            log.debug(f"  [LITIGANTES] {etiqueta}: pestaña no encontrada")
            return resultado
        tab_link.click()
        page.wait_for_selector(
            '#litigantesCiv tbody tr',
            state='visible',
            timeout=8000
        )

        # 2. Leer SOLO la tabla dentro de #litigantesCiv
        filas = page.locator('#litigantesCiv tbody tr').all()

        for fila in filas:
            celdas = fila.locator('td').all()
            if len(celdas) < 4:
                continue

            participante = celdas[0].inner_text().strip()
            nombre_raw = celdas[3].inner_text().strip()

            # Limpiar nombre: quitar "(Poder Amplio)" y espacios extra
            nombre = re.sub(r'\s*\(.*?\)\s*$', '', nombre_raw).strip()
            nombre = re.sub(r'\s+', ' ', nombre)

            if participante.startswith('DDO') and resultado['nombre_ddo'] is None:
                resultado['nombre_ddo'] = nombre
            elif participante.startswith('DTE') and resultado['nombre_dte'] is None:
                resultado['nombre_dte'] = nombre

        # Log de resultados
        if resultado['nombre_dte'] or resultado['nombre_ddo']:
            log.info(f"  ✓ DTE (OJV): {resultado['nombre_dte']}")
            log.info(f"  ✓ DDO (OJV): {resultado['nombre_ddo']}")
        else:
            log.warning(f"  [LITIGANTES] {etiqueta}: tabla visible pero sin DTE/DDO")

    except Exception as e:
        log.warning(f"  [LITIGANTES] {etiqueta}: {e}")

    # Siempre volver a pestaña Historia (necesaria para cuaderno/descarga)
    _volver_a_historia(page)
    return resultado


def _volver_a_historia(page):
    """Vuelve a la pestaña Historia tras consultar Litigantes."""
    try:
        tab_historia = page.locator('a[href="#702"]')
        if tab_historia.count() > 0:
            tab_historia.click()
            time.sleep(0.8)
            return
    except Exception:
        pass
    # Fallback: buscar por texto
    try:
        tab_historia = page.query_selector(
            'a:has-text("Historia"), '
            '[data-toggle="tab"]:has-text("Historia"), '
            'li a:has-text("Historia")'
        )
        if tab_historia:
            tab_historia.click()
            time.sleep(0.8)
    except Exception:
        pass


# ─────────────────────────────────────────────────────────────────
# Selección dinámica de cuaderno (por texto, no por posición)
# ─────────────────────────────────────────────────────────────────

def _seleccionar_cuaderno_dinamico(page, texto_cuaderno: str) -> bool:
    """
    Selecciona un cuaderno del dropdown #selCuaderno buscando por texto.
    Busca la opción cuyo texto contenga texto_cuaderno (case-insensitive).
    Después de seleccionar, espera a que la tabla de historia se recargue.
    Retorna True si se seleccionó, False si no se encontró.
    """
    try:
        page.wait_for_selector("#selCuaderno", timeout=15000)
    except Exception as e:
        log.warning(f"  Dropdown cuaderno no disponible: {e}")
        return False

    texto_lower = texto_cuaderno.lower()

    # Log de opciones disponibles para diagnóstico
    try:
        opciones = page.query_selector_all("#selCuaderno option")
        disponibles = [o.inner_text().strip() for o in opciones]
        log.info(f"  Cuadernos disponibles: {disponibles}")
    except Exception:
        pass

    # Esperar hasta 15 segundos a que carguen las opciones
    for _ in range(30):
        try:
            opciones = page.query_selector_all("#selCuaderno option")
            for opt in opciones:
                opt_texto = opt.inner_text().strip()
                if texto_lower in opt_texto.lower():
                    valor = opt.get_attribute("value")
                    page.select_option("#selCuaderno", value=valor)
                    log.info(f"  Cuaderno seleccionado: '{opt_texto}' (value={valor})")
                    # Esperar a que la tabla de historia se recargue tras cambio de cuaderno
                    time.sleep(2)
                    try:
                        page.wait_for_selector(
                            "#loadHistCuadernoCivil table tbody tr, "
                            "#historiaClv table tbody tr, "
                            "#modalDetalleCivil .table-responsive table tbody tr",
                            timeout=8000,
                        )
                    except Exception:
                        log.debug("  Tabla de historia no recargó tras cambio de cuaderno")
                    return True
        except Exception:
            pass
        time.sleep(0.5)

    # No encontrado: log de opciones disponibles para diagnóstico
    try:
        opciones = page.query_selector_all("#selCuaderno option")
        disponibles = [o.inner_text().strip() for o in opciones]
        log.warning(f"  Cuaderno '{texto_cuaderno}' no encontrado. Opciones: {disponibles}")
    except Exception:
        pass
    return False


# ─────────────────────────────────────────────────────────────────
# Función de procesamiento individual (wrapper sobre ojv_remates)
# ─────────────────────────────────────────────────────────────────

def _procesar_una_causa(page, context, causa: dict) -> dict:
    """
    Procesa una causa individual con Playwright.
    Devuelve la causa enriquecida con campos de resultado.
    """
    etiqueta = f"C-{causa['rol']}-{causa['año']}"
    log.info(f"{'='*55}")
    log.info(f"  Causa    : {etiqueta}")
    log.info(f"  Corte    : {causa.get('corte', '')}")
    log.info(f"  Tribunal : {causa.get('tribunal', '')}")
    log.info(f"{'='*55}")

    # Valores por defecto
    causa = {**causa,
             "tipo_procedimiento": "",
             "tipo_documento": "",
             "descargado": False,
             "ruta_pdf": "",
             "motivo_fallo": ""}

    if not causa.get("corte") or causa["corte"] == "DESCONOCIDA":
        log.warning(f"  {etiqueta}: M1 no pudo asignar Corte (falta en Excel o match < 80%)")
        causa["motivo_fallo"] = "M1: Corte DESCONOCIDA (revisar Excel)"
        return causa

    if not causa.get("tribunal"):
        log.warning(f"  {etiqueta}: M1 no extrajo tribunal del texto del PDF")
        causa["motivo_fallo"] = "M1: Sin tribunal en PDF"
        return causa

    # Limpiar formulario y buscar (la navegación inicial ya ocurrió en el loop externo)
    if not limpiar_formulario(page):
        causa["motivo_fallo"] = "M2: OJV timeout en formulario"
        return causa
    if not buscar_causa(page, causa["rol"], causa["año"],
                        causa["corte"], causa["tribunal"]):
        causa["motivo_fallo"] = "M2: Dropdown OJV rechazó tribunal (score < 85%)"
        return causa
    if not abrir_detalle(page, causa.get("rol"), causa.get("año")):
        causa["motivo_fallo"] = "OJV: causa no encontrada"
        return causa

    # Extraer nombres completos de DTE y DDO desde pestaña Litigantes
    litigantes = _extraer_litigantes_ojv(page, etiqueta)
    if litigantes['nombre_dte']:
        causa['demandante'] = litigantes['nombre_dte']
    if litigantes['nombre_ddo']:
        causa['demandado'] = litigantes['nombre_ddo']

    # Detectar tipo de procedimiento y filtrar los no aplicables
    # Procedimientos que NO sirven para inversión inmobiliaria
    PROCEDIMIENTOS_DESCARTADOS = [
        "liquidación simplificada",
        "liquidación concursal",
        "ordinario mayor cuantía",
        "ordinario menor cuantía",
        "ordinario mínima cuantía",
        "partición",
        "arbitral",
    ]

    es_ley_bancos = False
    es_ejecutivo_obligacion = False
    es_desposeimiento = False
    pudo_leer_modal = False
    procedimiento_detectado = ""
    try:
        modal = page.query_selector("#modalDetalleCivil, .modal.in, .modal.show")
        if modal:
            modal_texto = modal.inner_text().lower()
            pudo_leer_modal = True
            es_ley_bancos = "ley de bancos" in modal_texto
            es_ejecutivo_obligacion = ("ejecutivo" in modal_texto
                                       and "obligaci" in modal_texto)
            es_desposeimiento = "desposeimiento" in modal_texto

            # Extraer nombre del procedimiento para logging
            import re as _re
            m_proc = _re.search(r'proc\.?:\s*(.+?)(?:\n|$)', modal_texto)
            if m_proc:
                procedimiento_detectado = m_proc.group(1).strip()

            # Verificar contra lista explícita de procedimientos descartados
            for proc_desc in PROCEDIMIENTOS_DESCARTADOS:
                if proc_desc in modal_texto:
                    log.warning(f"  {etiqueta}: procedimiento descartado: '{procedimiento_detectado or proc_desc}'")
                    causa["motivo_fallo"] = f"procedimiento descartado: {procedimiento_detectado or proc_desc}"
                    _cerrar_modal(page)
                    return causa
    except Exception:
        pass

    if pudo_leer_modal and not es_ley_bancos and not es_ejecutivo_obligacion and not es_desposeimiento:
        log.warning(f"  {etiqueta}: procedimiento no aplicable: '{procedimiento_detectado}' — descartando")
        causa["motivo_fallo"] = f"procedimiento no aplicable: {procedimiento_detectado}"
        _cerrar_modal(page)
        return causa

    if es_ley_bancos:
        tipo_proc = "ley_bancos"
        cuaderno_objetivo = "Principal"
    elif es_desposeimiento:
        tipo_proc = "desposeimiento"
        cuaderno_objetivo = "Apremio"   # "Apremio de desposeimiento" contiene "Apremio"
    else:
        tipo_proc = "ejecutivo"
        cuaderno_objetivo = "Apremio"
    causa["tipo_procedimiento"] = tipo_proc
    log.info(f"  Proc.: {tipo_proc}")

    # Selección dinámica del cuaderno: buscar por texto, no por posición
    if not _seleccionar_cuaderno_dinamico(page, cuaderno_objetivo):
        log.warning(f"  Cuaderno '{cuaderno_objetivo}' no disponible para {etiqueta}")
        causa["motivo_fallo"] = f"OJV: cuaderno {cuaderno_objetivo} no encontrado"
        _cerrar_modal(page)
        return causa

    nombre_pdf = os.path.join(DESCARGAS_DIR, f"{etiqueta}_MANDAMIENTO.pdf")
    ok = False

    if es_ley_bancos:
        log.info(f"  [BASES DE REMATE]")
        ok = buscar_bases_remate(page, context, etiqueta)
        causa["tipo_documento"] = "bases_remate"
        nombre_pdf = os.path.join(DESCARGAS_DIR, f"{etiqueta}_BASES_REMATE.pdf")
    else:
        log.info(f"  [MANDAMIENTO]")
        ok = buscar_mandamiento(page, context, etiqueta)
        causa["tipo_documento"] = "mandamiento"
        nombre_pdf = os.path.join(DESCARGAS_DIR, f"{etiqueta}_MANDAMIENTO.pdf")

    if ok and os.path.exists(nombre_pdf):
        causa["descargado"] = True
        causa["ruta_pdf"] = nombre_pdf
        log.info(f"  ✓ Descargado: {os.path.basename(nombre_pdf)}")
    else:
        causa["motivo_fallo"] = "OJV: descarga fallida"
        log.warning(f"  ✗ No descargado: {etiqueta}")

    _cerrar_modal(page)
    return causa


def _cerrar_modal(page):
    """Cierra el modal de detalle."""
    for sel in ["button:has-text('Cerrar')", ".modal .close", "button.close"]:
        try:
            page.click(sel, timeout=2000)
            break
        except Exception:
            pass
    try:
        page.keyboard.press("Escape")
    except Exception:
        pass
    time.sleep(1)


# ─────────────────────────────────────────────────────────────────
# FUNCIÓN PÚBLICA — interface para el orquestador
# ─────────────────────────────────────────────────────────────────

def procesar_causas_ojv(causas: list[dict]) -> list[dict]:
    """
    Recibe la lista de causas del Módulo 1 y para cada una:
    1. Busca la causa en la OJV
    2. Detecta tipo de procedimiento (ejecutivo / ley de bancos)
    3. Descarga el documento correspondiente (mandamiento o bases de remate)
    4. Enriquece la causa con los campos de resultado

    Args:
        causas: lista de dicts del módulo 1

    Returns:
        Misma lista enriquecida con: tipo_procedimiento, tipo_documento,
        descargado (bool), ruta_pdf (str)
    """
    log.info(f"Iniciando Módulo 2 — {len(causas)} causa(s) a procesar")

    os.makedirs(DESCARGAS_DIR, exist_ok=True)

    # Filtrar causas sin corte o tribunal (no se puede buscar en OJV)
    causas_validas = [c for c in causas
                      if c.get("corte") and c["corte"] != "DESCONOCIDA"
                      and c.get("tribunal")]
    causas_invalidas = [c for c in causas if c not in causas_validas]

    log.info(f"  Procesables: {len(causas_validas)} | Sin corte/tribunal: {len(causas_invalidas)}")

    # Marcar las inválidas con campos vacíos
    for c in causas_invalidas:
        c.setdefault("tipo_procedimiento", "")
        c.setdefault("tipo_documento", "")
        c.setdefault("descargado", False)
        c.setdefault("ruta_pdf", "")
        c.setdefault("motivo_fallo", "OJV: tribunal no reconocido")

    resultados = list(causas_invalidas)

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False, slow_mo=100)
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()
        page.set_default_timeout(15000)

        if not navegar_a_consulta(page):
            log.error("No se pudo abrir el formulario OJV — abortando M2")
            browser.close()
            for c in causas_validas:
                c.setdefault("tipo_procedimiento", "")
                c.setdefault("tipo_documento", "")
                c.setdefault("descargado", False)
                c.setdefault("ruta_pdf", "")
                c.setdefault("motivo_fallo", "OJV: timeout")
                resultados.append(c)
            return resultados

        for i, causa in enumerate(causas_validas, 1):
            etiqueta = f"C-{causa['rol']}-{causa['año']}"
            log.info(f"\n[{i}/{len(causas_validas)}] {etiqueta}")

            # Filtro blacklist: causas con cuadernos restringidos u otros problemas conocidos
            if etiqueta in CAUSAS_IGNORADAS:
                log.info(f"  {etiqueta}: en CAUSAS_IGNORADAS — saltando")
                causa["tipo_procedimiento"] = ""
                causa["tipo_documento"] = ""
                causa["descargado"] = False
                causa["ruta_pdf"] = ""
                causa["motivo_fallo"] = "causa en blacklist (CAUSAS_IGNORADAS)"
                resultados.append(causa)
                continue

            try:
                causa_enriquecida = _procesar_una_causa(page, context, causa)
                resultados.append(causa_enriquecida)
            except KeyboardInterrupt:
                log.info("Detenido por el usuario")
                # Agregar causas restantes sin procesar
                for c in causas_validas[i:]:
                    c.setdefault("tipo_procedimiento", "")
                    c.setdefault("tipo_documento", "")
                    c.setdefault("descargado", False)
                    c.setdefault("ruta_pdf", "")
                    resultados.append(c)
                break
            except Exception as e:
                log.error(f"  ERROR en {etiqueta}: {e}")
                causa["tipo_procedimiento"] = ""
                causa["tipo_documento"] = ""
                causa["descargado"] = False
                causa["ruta_pdf"] = ""
                causa["motivo_fallo"] = f"OJV: error inesperado ({type(e).__name__}: {str(e)[:80]})"
                resultados.append(causa)
            time.sleep(2)

        browser.close()

    # Resumen
    descargados  = sum(1 for c in resultados if c.get("descargado"))
    ejecutivos   = sum(1 for c in resultados if c.get("tipo_procedimiento") == "ejecutivo")
    ley_bancos   = sum(1 for c in resultados if c.get("tipo_procedimiento") == "ley_bancos")
    desposeim    = sum(1 for c in resultados if c.get("tipo_procedimiento") == "desposeimiento")
    proc_no_apl  = sum(1 for c in resultados if c.get("motivo_fallo") == "procedimiento no aplicable")
    sin_corte    = len(causas_invalidas)

    log.info("=" * 55)
    log.info("Módulo 2 completado:")
    log.info(f"  Descargados exitosos      : {descargados}")
    log.info(f"  Ejecutivos                : {ejecutivos}")
    log.info(f"  Ley de Bancos             : {ley_bancos}")
    log.info(f"  Desposeimiento            : {desposeim}")
    log.info(f"  Proc. no aplicable        : {proc_no_apl}")
    log.info(f"  Sin corte/tribunal        : {sin_corte}")
    log.info(f"  ✓ TOTAL PROCESADAS        : {len(resultados)}")
    log.info("=" * 55)

    # Tabla detallada de causas sin descarga
    fallidos = [c for c in resultados if not c.get("descargado")]
    if fallidos:
        log.info("")
        log.info("=" * 90)
        log.info("  CAUSAS SIN DESCARGA — DETALLE DE FALLOS")
        log.info("=" * 90)
        log.info(f"  {'ROL':<14}| {'TRIBUNAL':<40}| MOTIVO")
        log.info(f"  {'-'*13}|{'-'*40}|{'-'*35}")
        for c in fallidos:
            rol_str = f"C-{c.get('rol','?')}-{c.get('año','?')}"
            tribunal = (c.get("tribunal") or c.get("tribunal_raw") or "?")[:38]
            motivo = c.get("motivo_fallo") or "desconocido"
            log.info(f"  {rol_str:<14}| {tribunal:<40}| {motivo}")
        log.info("=" * 90)

    return resultados


# ─────────────────────────────────────────────────────────────────
# Standalone: permite ejecutar modulo2_ojv.py directamente
# para probar con causas del módulo 1
# ─────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    from modulo1_parser import parsear_diarios

    print("Ejecutando Módulo 1 para obtener causas...")
    causas = parsear_diarios()
    print(f"Módulo 1: {len(causas)} causas")

    print("\nEjecutando Módulo 2 (OJV)...")
    causas = procesar_causas_ojv(causas)

    print(f"\n{'='*55}")
    print("RESULTADO MÓDULO 2:")
    for c in causas:
        estado = "✓" if c.get("descargado") else "✗"
        proc   = c.get("tipo_procedimiento", "?")
        doc    = c.get("tipo_documento", "?")
        print(f"  {estado} C-{c['rol']}-{c['año']} | {proc} | {doc}")
