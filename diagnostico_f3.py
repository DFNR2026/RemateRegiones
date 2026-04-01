"""
Diagnóstico pre-Filtro 3: analiza causas PENDIENTE_LIQUIDACION.

Uso:
    python diagnostico_f3.py
"""

import pandas as pd
from config import EXCEL_MADRE

# Leer pestaña _datos_internos
df = pd.read_excel(EXCEL_MADRE, sheet_name="_datos_internos", dtype=str)
df = df.fillna("")

print(f"Excel: {EXCEL_MADRE}")
print(f"Total filas: {len(df)}")
print(f"Columnas: {df.columns.tolist()}\n")

# Filtrar PENDIENTE_LIQUIDACION
mask = df["estado"] == "PENDIENTE_LIQUIDACION"
df_pl = df[mask].copy()
print(f"{'='*70}")
print(f"  CAUSAS PENDIENTE_LIQUIDACION: {len(df_pl)}")
print(f"{'='*70}\n")

if len(df_pl) == 0:
    print("No hay causas PENDIENTE_LIQUIDACION.")
    exit()

# Mostrar primeros 3 registros completos para verificar columnas
print("--- MUESTRA: 3 primeros registros completos ---")
for i, (_, row) in enumerate(df_pl.head(3).iterrows()):
    print(f"\n[{i+1}]")
    for col in df_pl.columns:
        val = row[col]
        if val and str(val).strip():
            print(f"  {col}: {val}")
print(f"\n{'='*70}\n")

# Detalle de cada causa
print("--- DETALLE POR CAUSA ---\n")
for _, row in df_pl.iterrows():
    rol = row.get("ROL", "?")
    año = row.get("AÑO", "?")
    tribunal = row.get("TRIBUNAL", "?")
    deuda = row.get("MONTO_DEUDA_CLP", "")
    monto_acta = row.get("monto_acta_remate", "")
    detalle = row.get("detalle_auditoria", "")
    log_dec = row.get("log_decision", "")
    notas = row.get("notas", "")

    # Calcular delta
    delta_str = ""
    try:
        d = int(str(deuda).replace(".", "").replace(",", "")) if deuda else 0
        a = int(str(monto_acta).replace(".", "").replace(",", "")) if monto_acta else 0
        if d > 0 and a > 0:
            delta = a - d
            delta_str = f"${delta:,}"
    except (ValueError, TypeError):
        pass

    # Detectar si tiene keyword de liquidación en detalle
    tiene_kw_liq = False
    KW_LIQ_MARKERS = ["senal 4:", "ordena liquidar", "liquidaci", "liquida el"]
    detalle_lower = (detalle + " " + log_dec + " " + notas).lower()
    for marker in KW_LIQ_MARKERS:
        if marker in detalle_lower:
            tiene_kw_liq = True
            break

    tiene_acta = bool(monto_acta and str(monto_acta).strip())

    print(f"C-{rol}-{año} | {tribunal}")
    print(f"  Deuda: {deuda}  |  Monto acta: {monto_acta}  |  Delta: {delta_str}")
    print(f"  KW liquidación: {'SI' if tiene_kw_liq else 'NO'}  |  Tiene acta: {'SI' if tiene_acta else 'NO'}")
    if detalle:
        print(f"  Detalle: {detalle[:120]}")
    if log_dec:
        print(f"  Log: {log_dec[:120]}")
    print()

# Resumen
print(f"{'='*70}")
print(f"  RESUMEN")
print(f"{'='*70}")

total = len(df_pl)
con_acta = 0
con_kw_liq = 0
con_ambas = 0

for _, row in df_pl.iterrows():
    monto_acta = row.get("monto_acta_remate", "")
    detalle = row.get("detalle_auditoria", "")
    log_dec = row.get("log_decision", "")
    notas = row.get("notas", "")

    tiene_acta = bool(monto_acta and str(monto_acta).strip())
    tiene_kw = False
    detalle_lower = (detalle + " " + log_dec + " " + notas).lower()
    for marker in KW_LIQ_MARKERS:
        if marker in detalle_lower:
            tiene_kw = True
            break

    if tiene_acta:
        con_acta += 1
    if tiene_kw:
        con_kw_liq += 1
    if tiene_acta and tiene_kw:
        con_ambas += 1

solo_acta = con_acta - con_ambas
solo_kw = con_kw_liq - con_ambas
sin_ninguna = total - con_acta - solo_kw

print(f"  Total PENDIENTE_LIQUIDACION:       {total}")
print(f"  Con monto_acta (pasaron F2):       {con_acta}")
print(f"  Con keyword liquidación (F1):      {con_kw_liq}")
print(f"  Con AMBAS (acta + liquidación):    {con_ambas}")
print(f"  Solo acta (sin kw liquidación):    {solo_acta}")
print(f"  Solo kw liquidación (sin acta):    {solo_kw}")
print(f"  Sin ninguna señal clara:           {sin_ninguna}")
print(f"{'='*70}")
