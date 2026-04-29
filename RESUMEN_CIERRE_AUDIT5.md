# RESUMEN DE CIERRE — Sesión Paso 0.5: refresh automático de MONTO_DEUDA_CLP

**Fecha de cierre:** 2026-04-29

Sesión enfocada en diagnosticar y resolver el bug de causas con `MONTO_DEUDA_CLP=$0` en el Excel madre del filtrador, e integrar la solución de forma permanente en `filtrador_saldos.py`.

---

## 1. Bug detectado

11 causas en estado `PENDIENTE_LIQUIDACION` con `MONTO_DEUDA_CLP=$0` en el Excel madre, a pesar de que los PDFs de mandamiento existían en `Descargas/`. Causa raíz: carrera de pipeline.

**Secuencia del bug:**
1. M2 falla al descargar PDFs (WAF, timeouts, errores OJV transitorios)
2. M3 salta esas causas porque `descargado=False`
3. Reporte M5 queda con deuda=$0
4. Filtrador importa las causas con $0
5. M2 sí descarga los PDFs en runs posteriores, pero **M3 nunca vuelve a correr** sobre ellas y M5 nunca regenera
6. El filtrador tampoco refresca el monto al hacer merge de reportes nuevos

Resultado: causas con PDF disponible y M3 capaz de extraer el monto, pero deuda registrada como $0 indefinidamente.

---

## 2. Diagnóstico

Inspección de los 10 PDFs problemáticos confirmó que **M3 funciona perfectamente sobre ellos**: 165 de 170 PDFs en `Descargas/` (97%) extraídos correctamente, incluyendo los 10 problemáticos. El bug no estaba en M3 sino en el handoff entre módulos.

Hallazgo adicional: `monto_acta_remate` se guarda como string en el filtrador (consistente con el resto de columnas numéricas que `_deuda_a_int` parsea al leer). En condiciones normales F2 calcula `_delta` correctamente; las 10 problemáticas tenían `_delta=None` porque cuando F2 corrió originalmente la deuda era $0 y abortó la comparación con `"Comparacion incompleta"`.

(C-2333-2024 era excepción legítima: vino de señal 4 "Ordena liquidar" → fue directo a F3 sin pasar por F2 acta-vs-deuda. Su `_delta=None` es correcto por flujo.)

---

## 3. Fix one-shot 1 — `fix_deuda_cero.py`

Script complementario, no integrado al filtrador. Función:
- Carga `Causas_posible_saldo.xlsx` (`_datos_internos`)
- Identifica causas con `MONTO_DEUDA_CLP` en {0, None, vacío} y estado activo
- Para cada candidata: busca PDF en `Descargas/C-{rol}-{año}_MANDAMIENTO.pdf`, ejecuta M3, actualiza fila
- Anexa marca `FixDeudaCero [fecha]` a `log_decision`

Diseño: dry-run por defecto, `--apply` para escribir, backup timestampeado, idempotente.

**Resultado:** 11/11 causas actualizadas. Verificadas contra montos esperados (10 conocidas + C-2333-2024 que apareció en el listado).

Archivado en `scripts_oneshot_archivados/` post-validación (commit `ed04e64`).

---

## 4. Fix one-shot 2 — `fix_recalcular_delta.py`

Complemento del fix anterior. El `fix_deuda_cero.py` actualizó `MONTO_DEUDA_CLP` pero no recalculó `_delta`. Script para 10 causas con `monto_acta` y `MONTO_DEUDA_CLP` ambos > 0 pero `_delta=None`.

CC verificó la fórmula real en `filtrador_saldos.py` L640-645 y replicó byte-a-byte los helpers `_deuda_a_int` (L324) y `_parsear_monto` (L309) para evitar acoplamiento al filtrador.

**Resultado:** 10/10 deltas calculados correctamente, escritos como int en el Excel. Deltas negativos (C-2120, C-3819, C-2729, C-1813) NO promovidos a `ELIMINADA` automáticamente — se respetó la regla de no cambiar estado, dejando que F3 confirme/descarte cuando aparezca el PDF de la liquidación.

Archivado en `scripts_oneshot_archivados/` post-validación.

---

## 5. Fix permanente — Paso 0.5 en `filtrador_saldos.py`

Para que el bug no vuelva a aparecer en runs futuros. Diff de 5 hunks (commit `f45ca6a`):

**Hunk 1:** Importar `DESCARGAS_DIR` desde config.

**Hunk 2:** Try-import de M3 con flag `_M3_DISPONIBLE`. Si M3 falla al importar (encoding, tessdata, deps), el filtrador sigue corriendo y el Paso 0.5 se autosaltea con log explicativo.

**Hunk 3:** Función `_paso_refresh_deuda_m3(df)`. Itera causas con deuda=$0 + estado activo, busca PDF en `Descargas/`, ejecuta M3, actualiza fila + log_decision con marca `RefreshDeudaM3`. Recálculo de `_delta` vía `_calcular_campos_derivados` (fórmula única, no duplicada). Persiste Excel mediante `_guardar_excel_formateado` (que internamente tiene `_guardar_excel_con_retry`). Errores M3 por causa no abortan el run, se acumulan en contador.

**Hunk 4:** Argparse `--skip-refresh-deuda` para debugging. Default: ejecutar.

**Hunk 5:** Invocación en orquestador entre Paso 0 (merge) y lanzamiento de workers. Gating explícito: corre en run normal, `--reaudit` y `--skip-pdf`. Salteado en `--solo-merge`, `--solo-filtro1`, `--solo-filtro3`, `--recheck-cargo`, `--recheck-liq`. El paso ejecuta una sola vez en el orquestador, NO en workers.

**Hotfix descubierto durante test:** `df.at[idx, "MONTO_DEUDA_CLP"] = monto_clp` falló porque pandas con `dtype=str` rechaza int directo. Cambio: `... = str(monto_clp)`. Consistente con el resto del filtrador donde todos los montos numéricos viven como string y `_deuda_a_int` parsea al leer.

---

## 6. Validación

**Run 1 (`--reaudit --workers 1 --skip-pdf`):**
- Paso 0.5: 77 evaluadas / 23 sin PDF / 4 M3 fallo / **50 actualizadas**
- Excel guardado, `_delta` recalculado vía `_calcular_campos_derivados`
- REAUDIT subsiguiente arrancó procesando 24 causas (vs 26 sin Paso 0.5) — efecto del refresh
- Crash con UnicodeEncodeError en `ojv_remates.py:207` durante REAUDIT (no Paso 0.5). Bug pre-existente en `_TeeWriter`, ver Sección 7.

**Run 2 (`--reaudit --workers 5 --skip-pdf`):**
- Paso 0.5: 27 evaluadas / 23 sin PDF / 4 M3 fallo / **0 actualizadas** → idempotencia confirmada
- Las 27 residuales son las mismas 23 sin PDF + 4 fallos de M3, estado estable
- REAUDIT clean en 2:07, los 5 workers exit 0, sin UnicodeEncodeError
- Hotfix string validado: muestras con `deuda` y `acta` como str, `_delta` como int correcto, fórmula `acta - deuda == delta` verificada

---

## 7. Deuda técnica anotada (no fixeada)

Bug pre-existente en `_TeeWriter.write()` (`filtrador_saldos.py:96`) cuando se ejecuta con `--workers 1` o cualquier modo que invoque OJV en el proceso principal:

- `ojv_remates.py` hace `print("✓ Formulario listo")` con caracteres U+2713 / U+2717
- `_TeeWriter` replica stdout a `sys.__stdout__` que en Windows es cp1252
- Falla con `UnicodeEncodeError`
- El log file (utf-8) sí escribe bien, por eso aparece OK en disco

**No bloquea producción** porque los runs reales son `--workers 5`+, donde los workers tienen `PYTHONIOENCODING=utf-8` heredado del Fix encoding del Audit4.

Fix futuro: ~3 líneas en `_TeeWriter.write` que sanitice o use `errors='replace'` al escribir a `sys.__stdout__`. PR separado cuando se aborde.

---

## 8. Convenciones aplicadas

- **No tocar M3** (`modulo3_extractor.py` intacto, solo importado).
- **No tocar otros módulos** (M1/M2/M5/main).
- **One-shots primero, integración después** — patrón que funcionó: validar la lógica con scripts standalone antes de tocar el filtrador.
- **Replicación byte-a-byte de helpers** (`_deuda_a_int`, `_parsear_monto`) en scripts one-shot para evitar acoplamiento.
- **Reutilización de `_calcular_campos_derivados`** en el Paso 0.5 — fórmula única en el código, herencia automática de cambios futuros.
- **Idempotencia obligatoria** en cada fix (re-ejecutar no rompe ni duplica nada).

---

## 9. Estado del repo al cierre

**Branch:** `main` sincronizada con `origin/main` tras el push final del ciclo.

**Commits del ciclo (orden cronológico):**

| Hash | Tipo | Descripción |
|---|---|---|
| `f45ca6a` | feat | Paso 0.5 — refresh MONTO_DEUDA_CLP via M3 sobre PDFs locales |
| `ed04e64` | chore | archivar scripts one-shot del ciclo deuda=$0 |
| `a681f5b` | chore | trackear `causas_eliminadas_historial.csv` (output de Fix B) |
| (este commit) | docs | `RESUMEN_CIERRE_AUDIT5.md` — este documento |

**Working tree pendiente al cierre** (no relacionado al ciclo, son outputs de los runs de validación que se regeneran):

- `Causas con liq/Causas_posible_saldo.xlsx` modificado (state post-Paso-0.5: 50 deudas refrescadas + REAUDIT del run 2)
- `Causas con liq/Causas_posible_saldo_29_abril_2026.xlsx` (copia fechada generada por `_guardar_excel_formateado`)

**Repositorio:** `github.com/DFNR2026/RemateRegiones`.
