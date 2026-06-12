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
import json
import logging
import argparse
import shutil
import subprocess

# Forzar UTF-8 en stdout/stderr para que los print() de ojv_remates.py
# (que usan ✓ ✗ → ⚠ etc.) no fallen en terminales Windows con cp1252.
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

from playwright.sync_api import sync_playwright
from playwright_stealth import Stealth

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
# Workers subprocess
# ─────────────────────────────────────────────────────────────────
_BASE_DIR = os.path.dirname(os.path.abspath(__file__))
_TEMP_WORKERS_DIR = os.path.join(_BASE_DIR, "temp_workers_m2")
_DEFAULT_WORKERS = 5
_MAX_WORKERS = 10


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
# Serialización para workers subprocess
# ─────────────────────────────────────────────────────────────────

def _serializar_causas(causas_con_idx):
    """Convierte lista de (idx, causa_dict) a JSON-serializable."""
    resultado = []
    for idx, causa in causas_con_idx:
        c = {"_idx": idx}
        for k, v in causa.items():
            if hasattr(v, "isoformat"):
                c[k] = v.isoformat()
            elif isinstance(v, float) and (v != v):  # NaN
                c[k] = ""
            else:
                c[k] = v
        resultado.append(c)
    return resultado


def _deserializar_causas(lista):
    """Convierte lista de dicts JSON a lista de (idx, causa_dict)."""
    resultado = []
    for c in lista:
        idx = c.pop("_idx")
        resultado.append((idx, c))
    return resultado


# ─────────────────────────────────────────────────────────────────
# Worker subprocess
# ─────────────────────────────────────────────────────────────────

def _run_worker_m2(args):
    """Proceso hijo: abre su propio Playwright, procesa chunk, escribe resultado."""
    worker_id = args.worker_id
    log.info(f"[Worker {worker_id}] Iniciando...")

    with open(args.chunk_file, "r", encoding="utf-8") as f:
        chunk_data = json.load(f)
    causas_chunk = _deserializar_causas(chunk_data)

    log.info(f"[Worker {worker_id}] {len(causas_chunk)} causas a procesar")

    resultados = {}

    with sync_playwright() as p:
        _profile_dir = os.path.join(_BASE_DIR, f".chrome-profile-w{worker_id}")
        context = p.chromium.launch_persistent_context(
            _profile_dir,
            headless=False,
            slow_mo=100,
            channel="chrome",
            args=["--disable-blink-features=AutomationControlled"],
            accept_downloads=True,
        )
        page = context.pages[0] if context.pages else context.new_page()
        Stealth().apply_stealth_sync(page)
        page.set_default_timeout(15000)

        if not navegar_a_consulta(page):
            log.error(f"[Worker {worker_id}] No se pudo navegar a OJV")
            for idx, causa in causas_chunk:
                causa.setdefault("tipo_procedimiento", "")
                causa.setdefault("tipo_documento", "")
                causa.setdefault("descargado", False)
                causa.setdefault("ruta_pdf", "")
                causa.setdefault("motivo_fallo", "OJV: timeout")
                resultados[idx] = causa
            try:
                context.close()
            except Exception:
                pass
        else:
            for i, (idx, causa) in enumerate(causas_chunk, 1):
                etiqueta = f"C-{causa['rol']}-{causa['año']}"
                log.info(f"[Worker {worker_id}] [{i}/{len(causas_chunk)}] {etiqueta}")

                if etiqueta in CAUSAS_IGNORADAS:
                    log.info(f"  {etiqueta}: en CAUSAS_IGNORADAS -- saltando")
                    causa["tipo_procedimiento"] = ""
                    causa["tipo_documento"] = ""
                    causa["descargado"] = False
                    causa["ruta_pdf"] = ""
                    causa["motivo_fallo"] = "causa en blacklist (CAUSAS_IGNORADAS)"
                    resultados[idx] = causa
                    continue

                try:
                    causa_enriquecida = _procesar_una_causa(page, context, causa)
                    resultados[idx] = causa_enriquecida
                except Exception as e:
                    log.error(f"  ERROR en {etiqueta}: {e}")
                    causa["tipo_procedimiento"] = ""
                    causa["tipo_documento"] = ""
                    causa["descargado"] = False
                    causa["ruta_pdf"] = ""
                    causa["motivo_fallo"] = f"OJV: error inesperado ({type(e).__name__}: {str(e)[:80]})"
                    resultados[idx] = causa
                time.sleep(2)

            try:
                context.close()
            except Exception:
                pass

    # Serializar resultados
    result_list = _serializar_causas([(idx, c) for idx, c in resultados.items()])
    with open(args.result_file, "w", encoding="utf-8") as f:
        json.dump(result_list, f, ensure_ascii=False)

    log.info(f"[Worker {worker_id}] Finalizado. {len(resultados)} causas procesadas.")


def _ejecutar_paralelo_m2(causas_validas, n_workers):
    """Lanza N subprocesos, reparte por round-robin, merge resultados."""
    # Limpiar perfiles Chrome de workers para evitar estado WAF cacheado
    for i in range(1, n_workers + 1):
        wp = os.path.join(_BASE_DIR, f".chrome-profile-w{i}")
        if os.path.exists(wp):
            shutil.rmtree(wp, ignore_errors=True)

    os.makedirs(_TEMP_WORKERS_DIR, exist_ok=True)

    # Round-robin distribution
    chunks = {w: [] for w in range(1, n_workers + 1)}
    for i, causa in enumerate(causas_validas):
        w_id = (i % n_workers) + 1
        chunks[w_id].append((i, causa))

    # Lanzar workers (escalonados 3s entre cada uno, excepto el ultimo)
    procesos = []
    chunks_activos = [(w_id, chunk) for w_id, chunk in chunks.items() if chunk]

    for i, (w_id, chunk) in enumerate(chunks_activos):
        chunk_file = os.path.join(_TEMP_WORKERS_DIR, f"chunk_{w_id}.json")
        result_file = os.path.join(_TEMP_WORKERS_DIR, f"result_{w_id}.json")

        with open(chunk_file, "w", encoding="utf-8") as f:
            json.dump(_serializar_causas(chunk), f, ensure_ascii=False)

        cmd = [
            sys.executable, os.path.join(_BASE_DIR, "modulo2_ojv.py"),
            "--worker-mode",
            "--chunk-file", chunk_file,
            "--result-file", result_file,
            "--worker-id", str(w_id),
        ]

        log.info(f"  Lanzando Worker {w_id} ({len(chunk)} causas)...")
        proc = subprocess.Popen(cmd, cwd=_BASE_DIR)
        procesos.append((w_id, proc, result_file))

        if i < len(chunks_activos) - 1:
            time.sleep(3)

    # Esperar a que terminen
    for w_id, proc, _ in procesos:
        proc.wait()
        log.info(f"  Worker {w_id} terminado (exit code: {proc.returncode})")

    # Merge resultados
    all_resultados = {}
    for w_id, _, result_file in procesos:
        if os.path.exists(result_file):
            with open(result_file, "r", encoding="utf-8") as f:
                data = json.load(f)
            for idx, causa in _deserializar_causas(data):
                all_resultados[idx] = causa
        else:
            log.warning(f"  Worker {w_id}: sin archivo de resultados")

    # Limpiar temporales
    for fname in os.listdir(_TEMP_WORKERS_DIR):
        try:
            os.remove(os.path.join(_TEMP_WORKERS_DIR, fname))
        except OSError:
            pass
    try:
        os.rmdir(_TEMP_WORKERS_DIR)
    except OSError:
        pass

    # Reconstruir lista ordenada
    return [all_resultados[i] for i in sorted(all_resultados)]


# ─────────────────────────────────────────────────────────────────
# FUNCIÓN PÚBLICA — interface para el orquestador
# ─────────────────────────────────────────────────────────────────

def procesar_causas_ojv(causas: list[dict], n_workers: int = 1) -> list[dict]:
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

    n_workers = min(max(n_workers, 1), _MAX_WORKERS)

    if n_workers > 1 and len(causas_validas) > 1:
        # Modo paralelo
        log.info(f"  Modo paralelo: {n_workers} workers para {len(causas_validas)} causas")
        resultados_workers = _ejecutar_paralelo_m2(causas_validas, n_workers)
        resultados.extend(resultados_workers)
    else:
        # Modo secuencial (original)
        # Limpiar perfil Chrome para evitar estado WAF cacheado
        _profile_dir = os.path.join(_BASE_DIR, ".chrome-profile")
        if os.path.exists(_profile_dir):
            shutil.rmtree(_profile_dir, ignore_errors=True)
        with sync_playwright() as p:
            _profile_dir = os.path.join(_BASE_DIR, ".chrome-profile")
            context = p.chromium.launch_persistent_context(
                _profile_dir,
                headless=False,
                slow_mo=100,
                channel="chrome",
                args=["--disable-blink-features=AutomationControlled"],
                accept_downloads=True,
            )
            page = context.pages[0] if context.pages else context.new_page()
            Stealth().apply_stealth_sync(page)
            page.set_default_timeout(15000)

            if not navegar_a_consulta(page):
                log.error("No se pudo abrir el formulario OJV -- abortando M2")
                try:
                    context.close()
                except Exception:
                    pass
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

                # Filtro blacklist
                if etiqueta in CAUSAS_IGNORADAS:
                    log.info(f"  {etiqueta}: en CAUSAS_IGNORADAS -- saltando")
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

            try:
                context.close()
            except Exception:
                pass

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

def _parse_worker_args():
    """Parse args cuando se invoca como worker subprocess."""
    parser = argparse.ArgumentParser()
    parser.add_argument("--worker-mode", action="store_true")
    parser.add_argument("--chunk-file", type=str)
    parser.add_argument("--result-file", type=str)
    parser.add_argument("--worker-id", type=int, default=0)
    return parser.parse_args()


if __name__ == "__main__":
    # Detectar si es invocacion como worker subprocess
    if "--worker-mode" in sys.argv:
        args = _parse_worker_args()
        _run_worker_m2(args)
        sys.exit(0)

    # Standalone testing (original)
    print("Modo standalone v1 deshabilitado (Tanda D). Usa los flags "
          "del módulo o el pipeline completo: python main.py --docx ...")
    raise SystemExit(0)
