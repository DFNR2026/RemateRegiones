"""
fix_ocr_actas.py
================
Reprocesa las actas escaneadas cuyo monto de adjudicacion no se extrajo
porque el OCR estaba roto (TESSDATA_PREFIX no seteado -> Tesseract no
encontraba spa.traineddata).

37 causas quedaron en PENDIENTE_LIQUIDACION con "monto_adj=$0 ...
Comparacion incompleta" en log_decision. El acta estaba descargada pero
el OCR fallaba en todas las paginas y devolvia texto vacio.

REQUIERE: el fix de TESSDATA_PREFIX en filtrador_saldos._ocr_pdf (Tarea 1)
YA APLICADO. Sin ese fix, el OCR seguira fallando y este script reportara
"OCR sin monto" para todas las candidatas.

USO:
    python fix_ocr_actas.py            # DRY-RUN (no guarda)
    python fix_ocr_actas.py --apply    # EJECUCION REAL (backup + escribe)

LOGICA por candidata (acta acta_{rol}_{ano}.pdf en ACTAS_DIR):
  - cargo al credito              -> estado=ELIMINADA (decision del sistema)
  - monto>0 y deuda>0 y monto<deuda -> estado=ELIMINADA (sin excedente)
  - monto>0 y monto>=deuda        -> mantiene PENDIENTE_LIQUIDACION,
                                     setea monto_acta_remate, recalcula _delta
  - OCR sin monto                 -> deja como esta (revisar PDF manualmente)

Idempotente: si log_decision ya tiene marca "FixOcrActas", saltea.
Montos como string (consistencia dtype=str del filtrador); _delta como int.

NOTA: importa filtrador_saldos para reusar _analizar_pdf_acta,
_calcular_campos_derivados y _guardar_excel_formateado (no se reescriben).
Eso dispara el setup de logging del filtrador; los logs de este script
salen por esos handlers. Mantener mensajes en ASCII.
"""
import argparse
import os
import shutil
import sys
import time
from datetime import datetime

import pandas as pd

from config import EXCEL_MADRE, ACTAS_DIR
# Helpers del filtrador (importar, no reescribir). El fix de TESSDATA_PREFIX
# viaja dentro de _ocr_pdf, que usa _analizar_pdf_acta.
from filtrador_saldos import (
    _analizar_pdf_acta,
    _calcular_campos_derivados,
    _guardar_excel_formateado,
    _deuda_a_int,
    _COLS_MADRE,
)

SHEET = "_datos_internos"
FECHA = "2026-05-24"
PREFIJO_LOG = f"FixOcrActas [{FECHA}]"
ESTADOS_EXCLUIDOS = {"ELIMINADA", "ELIMINAR", "EXCEDENTE_CONFIRMADO"}


def _log(msg):
    """Print solo ASCII (Windows cp1252)."""
    safe = msg.encode("ascii", errors="replace").decode("ascii")
    print(safe)


def _es_candidata_log(log_dec):
    """True si el log_decision indica acta sin monto extraido."""
    s = str(log_dec or "")
    return ("monto_adj=$0" in s) or ("Comparacion incompleta" in s)


def _cargar_df():
    """Carga _datos_internos como DataFrame (dtype=str, igual que el filtrador)."""
    df = pd.read_excel(EXCEL_MADRE, sheet_name=SHEET, dtype=str, engine="openpyxl")
    for col in _COLS_MADRE:
        if col not in df.columns:
            df[col] = ""
    return df


def _backup_excel():
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    base, ext = os.path.splitext(EXCEL_MADRE)
    dst = f"{base}_backup_fix_ocr_actas_{ts}{ext}"
    shutil.copy2(EXCEL_MADRE, dst)
    return dst


def main():
    parser = argparse.ArgumentParser(
        description="Reprocesa actas escaneadas sin monto (OCR roto). Dry-run por defecto.")
    parser.add_argument("--apply", action="store_true",
                        help="Aplica los cambios al Excel (sin esto corre en dry-run).")
    parser.add_argument("--limit", type=int, default=0,
                        help="Procesar solo las primeras N candidatas (0 = sin limite).")
    args = parser.parse_args()

    if not os.path.exists(EXCEL_MADRE):
        _log(f"ERROR: no existe {EXCEL_MADRE}")
        sys.exit(2)
    if not os.path.isdir(ACTAS_DIR):
        _log(f"ERROR: no existe ACTAS_DIR {ACTAS_DIR}")
        sys.exit(2)

    modo = "APPLY" if args.apply else "DRY-RUN"
    _log("=" * 70)
    _log(f"fix_ocr_actas.py  ({modo})  fecha: {FECHA}")
    _log(f"Excel:    {EXCEL_MADRE}")
    _log(f"ACTAS:    {ACTAS_DIR}")
    _log("=" * 70)

    df = _cargar_df()

    # ── Identificar candidatas ──
    candidatas = []
    for idx, row in df.iterrows():
        estado = str(row.get("estado", "") or "").strip()
        if estado in ESTADOS_EXCLUIDOS:
            continue
        if not _es_candidata_log(row.get("log_decision", "")):
            continue
        if _deuda_a_int(row.get("monto_acta_remate", "")) > 0:
            continue  # ya tiene monto
        if "FixOcrActas" in str(row.get("log_decision", "")):
            _log(f"  [IDEMPOTENT] fila {idx}: ya procesada por FixOcrActas, salteo")
            continue
        candidatas.append(idx)

    _log(f"\nCandidatas encontradas: {len(candidatas)}")
    if args.limit > 0:
        candidatas = candidatas[:args.limit]
        _log(f"Limitando a las primeras {len(candidatas)}")
    if not candidatas:
        _log("Nada que hacer. Salgo.")
        return

    # ── Procesar cada candidata ──
    # accion: 'eliminada_cargo' | 'eliminada_menor' | 'recuperado' | 'sin_monto' | 'sin_acta'
    resultados = []
    for idx in candidatas:
        rol = str(df.at[idx, "ROL"] or "").strip()
        ano = str(df.at[idx, "AÑO"] or "").strip()
        deuda = _deuda_a_int(df.at[idx, "MONTO_DEUDA_CLP"])
        filename = f"acta_{rol}_{ano}.pdf"
        filepath = os.path.join(ACTAS_DIR, filename)

        if not os.path.exists(filepath):
            _log(f"  [SIN ACTA] C-{rol}-{ano}: no existe {filename}")
            resultados.append((idx, rol, ano, "sin_acta", None, deuda))
            continue

        try:
            analisis = _analizar_pdf_acta(filepath)
        except Exception as e:
            _log(f"  [ERROR] C-{rol}-{ano}: fallo analizando acta: {e}")
            resultados.append((idx, rol, ano, "sin_monto", None, deuda))
            continue

        cargo = bool(analisis.get("cargo_al_credito"))
        monto = analisis.get("monto_adjudicacion")  # int | None

        if cargo:
            accion = "eliminada_cargo"
        elif monto and monto > 0:
            if deuda > 0 and monto < deuda:
                accion = "eliminada_menor"
            else:
                accion = "recuperado"
        else:
            accion = "sin_monto"

        resultados.append((idx, rol, ano, accion, monto, deuda))

        # Log por causa (dry-run y apply)
        monto_fmt = f"${monto:,}".replace(",", ".") if monto else "$0"
        deuda_fmt = f"${deuda:,}".replace(",", ".") if deuda else "$0"
        etiqueta = {
            "eliminada_cargo": "ELIMINADA (cargo al credito)",
            "eliminada_menor": f"ELIMINADA (monto {monto_fmt} < deuda {deuda_fmt})",
            "recuperado": f"RECUPERADO monto={monto_fmt} (>= deuda {deuda_fmt})",
            "sin_monto": "OCR sin monto (revisar PDF)",
        }[accion]
        _log(f"  [{'DRY' if not args.apply else 'APPLY'}] C-{rol}-{ano}: {etiqueta}")

    # ── Resumen ──
    from collections import Counter
    cnt = Counter(r[3] for r in resultados)
    _log("\n" + "=" * 70)
    _log("Resumen:")
    _log(f"  Candidatas:                 {len(resultados)}")
    _log(f"  Sin acta (PDF no existe):   {cnt.get('sin_acta', 0)}")
    _log(f"  ELIMINADA (cargo credito):  {cnt.get('eliminada_cargo', 0)}")
    _log(f"  ELIMINADA (monto < deuda):  {cnt.get('eliminada_menor', 0)}")
    _log(f"  Recuperadas (monto >= deuda): {cnt.get('recuperado', 0)}")
    _log(f"  OCR sin monto:              {cnt.get('sin_monto', 0)}")
    _log("=" * 70)

    if not args.apply:
        _log("\nDRY-RUN: cambios NO guardados. Re-correr con --apply para escribir.")
        return

    # ── Aplicar al DataFrame ──
    n_cambios = 0
    for idx, rol, ano, accion, monto, deuda in resultados:
        log_actual = str(df.at[idx, "log_decision"] or "")

        if accion == "eliminada_cargo":
            df.at[idx, "estado"] = "ELIMINADA"
            nueva = f"{PREFIJO_LOG}: cargo al credito -> ELIMINADA"
        elif accion == "eliminada_menor":
            df.at[idx, "estado"] = "ELIMINADA"
            nueva = (f"{PREFIJO_LOG}: monto ${monto:,} < deuda ${deuda:,} -> ELIMINADA"
                     ).replace(",", ".")
        elif accion == "recuperado":
            df.at[idx, "monto_acta_remate"] = str(monto)
            nueva = (f"{PREFIJO_LOG}: monto ${monto:,} >= deuda, delta recalculado"
                     ).replace(",", ".")
        else:  # sin_monto / sin_acta
            nueva = f"{PREFIJO_LOG}: OCR sin monto (revisar PDF)"

        df.at[idx, "log_decision"] = f"{nueva}\n{log_actual}".strip()
        n_cambios += 1

    # Recalcular _delta con la formula unica del filtrador
    df = _calcular_campos_derivados(df)

    _log(f"\nFilas modificadas en memoria: {n_cambios}")
    _log("Backup del Excel antes de escribir...")
    bak = _backup_excel()
    _log(f"  Backup: {os.path.basename(bak)}")

    _log("Guardando Excel (via _guardar_excel_formateado)...")
    _guardar_excel_formateado(df)
    _log("Excel guardado OK.")


if __name__ == "__main__":
    main()
