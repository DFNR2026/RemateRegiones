"""
fix_rescate_actas.py
====================
One-shot: rescata actas de remate "escondidas" bajo un nombre de tramite
generico ("Actuacion", "Mero tramite", "Certificado") en la tabla de OJV.

PROBLEMA: ~19 causas quedan en PENDIENTE_LIQUIDACION/PENDIENTE_ACTA sin
monto_acta_remate ni _delta porque _evaluar_filtros_tabla solo busca el
literal "acta de remate" en desc_tramite. El acta existe pero bajo otro
nombre, asi que no se detecta.

SOLUCION: para cada causa atascada, abre OJV, lee la tabla completa, busca
en una ventana de fechas alrededor del remate las filas con tramite
generico (lista blanca, sin ruido de la lista negra), descarga cada
candidata y la analiza con _analizar_pdf_acta (OCR ya arreglado en Audit6.1).
Si es acta -> aplica la misma decision que el reaudit (ELIMINADA / delta).

USO:
    python fix_rescate_actas.py            # DRY-RUN (no escribe)
    python fix_rescate_actas.py --apply    # backup + escribe Excel

REQUIERE: fix de TESSDATA_PREFIX en _ocr_pdf (Audit6.1) ya aplicado.
NO integra nada al filtrador; es one-shot de validacion.
"""
# reconfigure ANTES de importar filtrador_saldos: este envuelve sys.stdout con
# _TeeWriter(sys.__stdout__, ...). Si dejamos el stream en cp1252, los print()
# de ojv_remates con tildes/simbolos rompen con UnicodeEncodeError.
import sys
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

import argparse
import os
import shutil
import time
import unicodedata
from datetime import date, datetime, timedelta

import pandas as pd
from playwright.sync_api import sync_playwright
from playwright_stealth import Stealth

from config import EXCEL_MADRE, DESCARGAS_DIR, BASE_DIR
from ojv_remates import (
    navegar_a_consulta,
    buscar_causa,
    abrir_detalle,
    limpiar_formulario,
)
from filtrador_saldos import (
    _leer_tabla_historial_completa,
    _encontrar_fila_por_folio,
    _descargar_pdf_desde_fila,
    _analizar_pdf_acta,
    _calcular_campos_derivados,
    _guardar_excel_formateado,
    _deuda_a_int,
    _buscar_causa_fallback_todos,
    _cerrar_modal_detalle,
    KEYWORDS_LIQUIDACION,
    _COLS_MADRE,
)

SHEET = "_datos_internos"
FECHA = date.today().isoformat()
PREFIJO_LOG = f"RescateActa [{FECHA}]"
ESTADOS_OBJETIVO = {"PENDIENTE_ACTA", "PENDIENTE_LIQUIDACION"}

RESCATE_DIR = os.path.join(DESCARGAS_DIR, "rescate_actas")
PROFILE_DIR = os.path.join(BASE_DIR, ".chrome-profile-rescate")

VENTANA_PRE_DIAS = 2
VENTANA_POST_DIAS = 14

# Lista blanca: tramite generico bajo el que se esconde el acta (contains, normalizado).
LISTA_BLANCA = ["actuacion", "mero tramite", "certificado"]
# Lista negra: si el desc contiene cualquiera, NO es candidata (anti-ruido).
LISTA_NEGRA = [
    "no postores", "bases", "fija", "suspend", "reprograma", "liquidaci",
    "ordena liquidar", "solicita liquidaci", "pone en conocimiento", "mandamiento",
    "requerimiento", "notificaci", "exhort", "oficio", "gir", "escritura",
    "aprueba", "propone", "modifica", "acompan", "curso progresivo", "certifiquese",
    "no ha lugar", "traslado", "citacion", "avenimiento", "transaccion",
    "conciliacion", "da cuenta de pago", "alzamiento", "inscripcion", "embargo",
    "nulo", "publicaci",
]


def _log(msg):
    """Print ASCII-safe (defensivo; stdout ya esta en utf-8)."""
    print(str(msg).encode("ascii", errors="replace").decode("ascii"))


def _vacio(v):
    """True si la celda esta vacia. Pandas dtype=str lee celdas vacias como
    float('nan'); str(nan)='nan' es truthy, por eso hay que mapear nan/none/''."""
    return str(v or "").strip().lower() in ("", "nan", "none")


def _norm(s):
    """minuscula + sin tildes."""
    s = str(s or "").lower().strip()
    s = unicodedata.normalize("NFKD", s)
    return "".join(c for c in s if not unicodedata.combining(c))


def _es_blanca(desc):
    nd = _norm(desc)
    if nd == "":
        return True
    return any(w in nd for w in LISTA_BLANCA)


def _es_negra(desc):
    nd = _norm(desc)
    return any(b in nd for b in LISTA_NEGRA)


def _es_nulo(desc):
    """Tramite de expediente anulado: prefijo [Nulo] en desc_tramite.
    No es acta valida (verificado: _norm preserva corchetes).
    """
    return "[nulo]" in _norm(desc)


def _matchea_liquidacion(desc):  # reservada (ventana ahora anclada al remate; ver _procesar_causa)
    nd = _norm(desc)
    return any(_norm(kw) in nd for kw in KEYWORDS_LIQUIDACION)


def _fmt(n):
    try:
        return f"${int(n):,}".replace(",", ".")
    except (ValueError, TypeError):
        return "-"


def _cargar_df():
    df = pd.read_excel(EXCEL_MADRE, sheet_name=SHEET, dtype=str, engine="openpyxl")
    for col in _COLS_MADRE:
        if col not in df.columns:
            df[col] = ""
    return df


def _backup_excel():
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    base, ext = os.path.splitext(EXCEL_MADRE)
    dst = f"{base}_backup_fix_rescate_actas_{ts}{ext}"
    shutil.copy2(EXCEL_MADRE, dst)
    return dst


def _buscar_con_retry(page, rol, anio, corte, tribunal, causa):
    """Mismo patron que el reaudit (3 intentos + fallback_todos)."""
    found = False
    for intento in range(1, 4):
        try:
            found = buscar_causa(page, rol, anio, corte, tribunal)
            if found:
                break
        except Exception as e:
            _log(f"    intento {intento}/3 buscar_causa fallo: {e}")
            if intento < 3:
                page.wait_for_timeout(2000)
    if not found:
        try:
            limpiar_formulario(page)
        except Exception:
            pass
        found = _buscar_causa_fallback_todos(page, rol, anio, corte, causa)
    return found


def _procesar_causa(page, idx, row):
    """Devuelve dict con el resultado propuesto para una causa.

    keys: idx, rol, anio, tribunal, fecha_remate, folio (o None), desc,
          cargo, monto_adj, deuda, estado_actual, estado_prop, monto_acta,
          delta, accion, log_nuevo (o None), motivo
    """
    rol = str(row.get("ROL", "") or "").strip()
    anio = str(row.get("AÑO", "") or "").strip()
    corte = str(row.get("CORTE", "") or "").strip()
    tribunal = str(row.get("TRIBUNAL", "") or "").strip()
    estado_actual = str(row.get("estado", "") or "").strip()
    deuda = _deuda_a_int(row.get("MONTO_DEUDA_CLP", "0"))

    base = {
        "idx": idx, "rol": rol, "anio": anio, "tribunal": tribunal,
        "fecha_remate": str(row.get("fecha_remate", "") or ""),
        "folio": None, "desc": "", "cargo": False, "monto_adj": 0,
        "deuda": deuda, "estado_actual": estado_actual,
        "estado_prop": estado_actual, "monto_acta": None, "delta": None,
        "accion": "error_ojv", "log_nuevo": None, "motivo": "",
        "n_cand": 0, "n_analizadas": 0,
    }

    causa_dict = {"ROL": rol, "AÑO": anio, "CORTE": corte, "TRIBUNAL": tribunal}

    _cerrar_modal_detalle(page)
    _log(f"  [RESCATE] C-{rol}-{anio} ({tribunal})...")

    if not _buscar_con_retry(page, rol, anio, corte, tribunal, causa_dict):
        base["motivo"] = "no encontrada en OJV"
        return base
    try:
        if not abrir_detalle(page, rol, anio):
            base["motivo"] = "no se pudo abrir detalle"
            return base
    except Exception as e:
        base["motivo"] = f"abrir_detalle fallo: {e}"
        return base

    try:
        filas = _leer_tabla_historial_completa(page)
    except Exception as e:
        base["motivo"] = f"leer tabla fallo: {e}"
        return base
    if not filas:
        base["motivo"] = "tabla vacia"
        return base

    # Parsear fecha_remate
    fr_str = str(row.get("fecha_remate", "") or "").strip()
    try:
        fecha_remate = date.fromisoformat(fr_str)
    except (ValueError, TypeError):
        base["accion"] = "sin_fecha"
        base["motivo"] = "sin fecha_remate, no evaluable"
        return base

    # Ventana: SIEMPRE anclada al remate. El acta vive cerca de la fecha de
    # remate, NO de la senal de liquidacion (que puede ser de un ciclo anterior
    # o ser previa al acta). Ver C-2622 (acta 1 dia post-remate, posterior a la
    # senal) y C-4163 (senal de liquidacion anterior al remate actual).
    ventana_inicio = fecha_remate - timedelta(days=VENTANA_PRE_DIAS)
    ventana_fin = fecha_remate + timedelta(days=VENTANA_POST_DIAS)

    # Candidatas: en ventana + lista blanca + no lista negra
    candidatas = [
        f for f in filas
        if f.get("fecha_tramite")
        and ventana_inicio <= f["fecha_tramite"] <= ventana_fin
        and not _es_nulo(f.get("desc_tramite"))
        and _es_blanca(f.get("desc_tramite"))
        and not _es_negra(f.get("desc_tramite"))
    ]
    candidatas.sort(key=lambda f: (f["fecha_tramite"], f["folio"]))
    base["n_cand"] = len(candidatas)

    if not candidatas:
        base["accion"] = "sin_acta_gating"
        base["motivo"] = "0 candidatas en ventana (ninguna fila paso la lista blanca)"
        return base

    # Probar cada candidata hasta encontrar el acta
    for cand in candidatas:
        folio = cand["folio"]
        fila_tr = _encontrar_fila_por_folio(page, folio)
        if fila_tr is None:
            continue
        os.makedirs(RESCATE_DIR, exist_ok=True)
        pdf = _descargar_pdf_desde_fila(
            page, fila_tr, RESCATE_DIR, f"rescate_{rol}_{anio}_f{folio}.pdf"
        )
        if not pdf:
            continue
        base["n_analizadas"] += 1
        analisis = _analizar_pdf_acta(pdf)
        cargo = bool(analisis.get("cargo_al_credito"))
        monto_adj = analisis.get("monto_adjudicacion") or 0
        adj = analisis.get("adjudicatario", "indeterminado")
        if not (cargo or monto_adj):
            continue  # no es acta, probar siguiente

        # === ES ACTA: ruteo de 3 vias (Audit11) ===
        base["folio"] = folio
        base["desc"] = str(cand.get("desc_tramite", "") or "")
        base["cargo"] = cargo
        base["monto_adj"] = monto_adj

        if adj == "ejecutante":
            base["estado_prop"] = "ELIMINADA"
            base["accion"] = "eliminada_cargo"
            base["log_nuevo"] = (f"{PREFIJO_LOG}: Ejecutante adjudico (folio {folio})."
                                 " Sin excedente.")
        elif adj == "tercero":
            if monto_adj > 0:
                base["monto_acta"] = str(monto_adj)
            if monto_adj > 0 and deuda > 0 and monto_adj < deuda:
                base["estado_prop"] = "ELIMINADA"
                base["accion"] = "eliminada_menor"
                base["log_nuevo"] = (f"{PREFIJO_LOG}: Tercero adjudico {_fmt(monto_adj)}"
                                     f" < deuda {_fmt(deuda)} (folio {folio}).")
            elif monto_adj > 0 and deuda > 0:
                base["estado_prop"] = "PENDIENTE_LIQUIDACION"
                base["accion"] = "recuperado"
                base["delta"] = monto_adj - deuda
                base["log_nuevo"] = (f"{PREFIJO_LOG}: Tercero adjudico {_fmt(monto_adj)}"
                                     f" >= deuda {_fmt(deuda)} (folio {folio})."
                                     " Posible excedente.")
            else:
                base["estado_prop"] = "PENDIENTE_LIQUIDACION"
                base["accion"] = "incompleta"
                base["log_nuevo"] = (f"{PREFIJO_LOG}: Tercero, acta descargada, monto={_fmt(monto_adj)},"
                                     f" deuda={_fmt(deuda)} (folio {folio})."
                                     " Comparacion incompleta.")
        else:
            # adj == "indeterminado"
            base["estado_prop"] = "PENDIENTE_REVISION_MANUAL"
            base["accion"] = "revision_adjudicatario"
            base["log_nuevo"] = (f"{PREFIJO_LOG}: adjudicatario no determinado,"
                                 f" revisar manual (folio {folio}).")
        return base

    base["accion"] = "sin_acta_timing"
    base["motivo"] = (f"escaneo {base['n_cand']} candidatas en ventana"
                      f" ({base['n_analizadas']} analizadas), ninguna resulto acta")
    return base


def _imprimir_dryrun(resultados):
    _log("")
    _log("=" * 110)
    _log("DRY-RUN - cambios propuestos (NO escritos)")
    _log("=" * 110)
    hdr = (f"{'CAUSA':<14}{'FOLIO':<7}{'CARGO':<6}{'MONTO':<16}"
           f"{'DEUDA':<16}{'DELTA':<16}{'ESTADO':<22}")
    _log(hdr)
    _log("-" * 110)
    for r in resultados:
        causa = f"C-{r['rol']}-{r['anio']}"
        folio = str(r["folio"]) if r["folio"] is not None else "-"
        cargo = "si" if r["cargo"] else ("no" if r["folio"] is not None else "-")
        monto = _fmt(r["monto_adj"]) if r["monto_adj"] else "-"
        deuda = _fmt(r["deuda"]) if r["deuda"] else "-"
        delta = _fmt(r["delta"]) if r["delta"] is not None else "-"
        est = f"{r['estado_actual']}->{r['estado_prop']}" if r["estado_prop"] != r["estado_actual"] else r["estado_actual"]
        _log(f"{causa:<14}{folio:<7}{cargo:<6}{monto:<16}{deuda:<16}{delta:<16}{est:<22}")
        if r["folio"] is None:
            _log(f"    motivo: {r['motivo']}  [ventana: {r['n_cand']} cand, {r['n_analizadas']} analizadas]")
    _log("-" * 110)


def _imprimir_resumen(resultados):
    from collections import Counter
    c = Counter(r["accion"] for r in resultados)
    _log("")
    _log("Resumen:")
    _log(f"  Causas evaluadas:                       {len(resultados)}")
    _log(f"  ELIMINADA (cargo al credito):           {c.get('eliminada_cargo', 0)}")
    _log(f"  ELIMINADA (monto < deuda):              {c.get('eliminada_menor', 0)}")
    _log(f"  Recuperadas (delta posible):            {c.get('recuperado', 0)}")
    _log(f"  Acta sin comparacion (incompl):         {c.get('incompleta', 0)}")
    _log(f"  Sin acta - TIMING (hubo cand, 0 acta):  {c.get('sin_acta_timing', 0)}")
    _log(f"  Sin acta - GATING (0 cand en ventana):  {c.get('sin_acta_gating', 0)}")
    _log(f"  Sin fecha_remate:                       {c.get('sin_fecha', 0)}")
    _log(f"  Error OJV (no encontrada/tabla/etc):    {c.get('error_ojv', 0)}")


def main():
    parser = argparse.ArgumentParser(
        description="Rescata actas escondidas bajo tramites genericos. Dry-run por defecto.")
    parser.add_argument("--apply", action="store_true",
                        help="Aplica cambios al Excel (sin esto corre en dry-run).")
    parser.add_argument("--limit", type=int, default=0,
                        help="Procesar solo las primeras N candidatas (0 = sin limite).")
    args = parser.parse_args()

    if not os.path.exists(EXCEL_MADRE):
        _log(f"ERROR: no existe {EXCEL_MADRE}")
        sys.exit(2)

    modo = "APPLY" if args.apply else "DRY-RUN"
    _log("=" * 70)
    _log(f"fix_rescate_actas.py  ({modo})  fecha: {FECHA}")
    _log("=" * 70)

    df = _cargar_df()

    # === DIAGNOSTICO (temporal) - justo antes del filtro de seleccion ===
    from collections import Counter as _Counter
    _log("")
    _log("--- DIAGNOSTICO seleccion ---")
    _log(f"Excel (abs):    {os.path.abspath(EXCEL_MADRE)}")
    _log(f"Hoja leida:     {SHEET}")
    _log(f"Filas cargadas: {len(df)}")
    _cnt_estado = _Counter(str(e or "").strip() for e in df.get("estado", []))
    _log("Conteo por estado:")
    for _est, _n in sorted(_cnt_estado.items(), key=lambda x: -x[1]):
        _log(f"    {_est!r}: {_n}")
    _en_obj = _fail_monto = _fail_obs = _fail_marca = 0
    for _, _row in df.iterrows():
        if str(_row.get("estado", "") or "").strip() not in ESTADOS_OBJETIVO:
            continue
        _en_obj += 1
        if _deuda_a_int(_row.get("monto_acta_remate", "")) > 0:
            _fail_monto += 1
        if not _vacio(_row.get("observacion_abogado")):
            _fail_obs += 1
        if "RescateActa" in str(_row.get("log_decision", "") or ""):
            _fail_marca += 1
    _log(f"Filas en {sorted(ESTADOS_OBJETIVO)}: {_en_obj}")
    _log(f"  fallan por monto_acta no vacio:  {_fail_monto}")
    _log(f"  fallan por observacion no vacia: {_fail_obs}")
    _log(f"  fallan por marca RescateActa:    {_fail_marca}")
    _log("--- fin diagnostico ---")
    _log("")

    # Poblacion objetivo
    candidatas_idx = []
    for idx, row in df.iterrows():
        estado = str(row.get("estado", "") or "").strip()
        if estado not in ESTADOS_OBJETIVO:
            continue
        if _deuda_a_int(row.get("monto_acta_remate", "")) > 0:
            continue
        if not _vacio(row.get("observacion_abogado")):
            continue  # ya revisada por el abogado
        if "RescateActa" in str(row.get("log_decision", "") or ""):
            continue  # idempotencia
        candidatas_idx.append(idx)

    _log(f"\nCausas atascadas seleccionadas: {len(candidatas_idx)}")
    if args.limit > 0:
        candidatas_idx = candidatas_idx[:args.limit]
        _log(f"Limitando a las primeras {len(candidatas_idx)}")
    if not candidatas_idx:
        _log("Nada que hacer. Salgo.")
        return

    # Limpiar perfil dedicado
    if os.path.isdir(PROFILE_DIR):
        shutil.rmtree(PROFILE_DIR, ignore_errors=True)

    resultados = []
    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            PROFILE_DIR,
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
            _log("ERROR: no se pudo navegar a OJV. Abortando.")
            context.close()
            sys.exit(3)

        for idx in candidatas_idx:
            try:
                r = _procesar_causa(page, idx, df.loc[idx])
            except Exception as e:
                _log(f"  [ERROR] fila {idx}: {e}")
                r = None
            if r is not None:
                resultados.append(r)

        context.close()

    # Salida
    _imprimir_dryrun(resultados)
    _imprimir_resumen(resultados)

    if not args.apply:
        _log("\nDRY-RUN: cambios NO guardados. Re-correr con --apply para escribir.")
        return

    # Aplicar: solo las causas resueltas (folio encontrado + log_nuevo)
    resueltas = [r for r in resultados if r["log_nuevo"]]
    if not resueltas:
        _log("\nNo hay causas resueltas para aplicar.")
        return

    _log(f"\nAplicando {len(resueltas)} cambios. Backup del Excel...")
    bak = _backup_excel()
    _log(f"  Backup: {os.path.basename(bak)}")

    for r in resueltas:
        idx = r["idx"]
        df.at[idx, "estado"] = r["estado_prop"]
        if r["monto_acta"] is not None:
            df.at[idx, "monto_acta_remate"] = r["monto_acta"]
        log_actual = str(df.at[idx, "log_decision"] or "")
        df.at[idx, "log_decision"] = f"{r['log_nuevo']}\n{log_actual}".strip()

    # Recalcular _delta con la formula unica del filtrador
    df = _calcular_campos_derivados(df)

    _log("Guardando Excel (via _guardar_excel_formateado)...")
    _guardar_excel_formateado(df)
    _log("Excel guardado OK.")


if __name__ == "__main__":
    main()
