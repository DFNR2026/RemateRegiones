# RESUMEN DE CIERRE — AUDIT10: Cierre Frente A (discriminador 3 vias + QA eliminadas)
Fecha: 2026-06-05

## C-1855-2025 — RESUELTA
Estado PENDIENTE_LIQUIDACION. monto_acta $197.000.000, _delta $127.161.163 (folio 52,
adjudicada al TERCERO Ana Maria Moreno Gomez). Estaba a salvo desde el rescate del 27-may;
la premisa "radiactiva" de los Audits 7-9 quedo obsoleta (documentacion atrasada).

## Discriminador de adjudicatario de 3 vias (analisis_pdf.py)
Veredictos: ejecutante / tercero / indeterminado. Ancla en `\bse\s+adjudic\w*` (salta la
clausula condicional de las bases, que usa "adjudicarse"). cargo_al_credito = (adj != "tercero")
por compatibilidad con F2. Regla dura: indeterminado -> revision manual, NUNCA eliminar.

## QA de eliminadas (solo lectura, por folio exacto + gate de auto-verificacion)
11 causas ELIMINADA; 6 por "cargo". Re-analizadas con el PDF del folio exacto del log_decision.
Gate: las 4 calibradas (C-153/C-234/C-3459/C-4163) reprodujeron ejecutante + montos
216M/82M/109M/66.7M. Veredicto: 0 terceros eliminados por error. Las 6 son ejecutante con
cargo real. C-234 confirmada por escrito del Banco de Chile (capital demandado > adjudicacion;
el $1.3M del Excel era saldo parcial desfasado, sin remanente en efectivo -> excedente cero).

## Estandar de robustez
Ante adjudicatario indeterminado NUNCA se elimina (revision manual). El abogado es red para
falsos excedentes; nada es red para una eliminacion irreversible.