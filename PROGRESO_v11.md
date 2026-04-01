# PROGRESO — Sistema de Análisis de Remates Judiciales Chile

Última actualización: 2026-03-31 (F3 completado + 14 causas eliminadas por revisión. 2 excedentes confirmados, ~$84.4M total)

---

## Archivos del proyecto

### `config.py`
Configuración global del proyecto: rutas, claves API, constantes.
- `DIARIOS_DIR`, `DESCARGAS_DIR`, `CAUSAS_XLSX`
- `ANTHROPIC_API_KEY` (Claude API para M1)
- `DEMANDANTES_EXCLUIDOS`, `CORTES_RM`
- `CAUSAS_IGNORADAS` — blacklist de causas con cuadernos restringidos/inaccesibles en OJV
- `CAUSAS_LIQ_DIR`, `EXCEL_MADRE`, `LIQUIDACIONES_DIR` — rutas del Filtrador de Saldos
- `LOGS_LIQUI_DIR` — ruta de logs del Filtrador
- `EXCEL_LIQUIDACIONES` — ruta de Causas_con_liquidacion.xlsx
- `ACTAS_DIR` — ruta de PDFs de actas descargados por F2

### `modulo1_parser.py` ✅ COMPLETADO (v1 — respaldo)
Parser original de avisos de remate. Usa Claude Sonnet 4.6 para TODOS los bloques.
Mantenido como fallback. **En producción se usa `v2_experimental/modulo1_v2.py`.**

### `v2_experimental/modulo1_v2.py` ✅ EN PRODUCCIÓN (desde 2026-03-20)
Parser v2 con estrategia "Regex-Flow Haiku": regex extrae tribunal y dirección primero, Haiku solo como fallback.

**Resultados run producción v2 (2026-03-21):**
- 196 bloques → 55 causas nuevas, 64 llamadas Haiku, ~$0.15, ~20 min

### `modulo2_ojv.py` ✅ COMPLETADO
### `modulo3_extractor.py` ✅ COMPLETADO
### `modulo5_reporte.py` ✅ COMPLETADO
### `main.py` ✅ COMPLETADO
### `ojv_remates.py` (v10.0+) — MOTOR OJV BASE — NO reescribir desde cero
### `causas_ojv.xlsx` — BD interna: REFERENCIA (235 tribunales) + CAUSAS (historial)

### `filtrador_saldos.bat`
Batch para ejecución semanal. Corre filtrador_saldos.py --workers 5 (merge+F1+F2+F3).

### `fix_eliminar_lote.py`
Script one-shot para eliminar 14 causas del Excel madre (8 revisión abogados + 6 verificación manual).
Causas notables eliminadas: C-3323-2024 (era excedente $49.6M, sin futuro comercial),
C-93-2024 (era excedente $4.8M, cargo al crédito FactorOne). Usa openpyxl directo.

### `Causas_con_liquidacion.xlsx`
Excel de salida de F3. Pestañas "Excedentes Confirmados" y "Revisión Manual".
Formato profesional para abogados. Separadores por fecha de ejecución para
identificar liquidaciones nuevas de cada run.

---

## Filtrador de Saldos — Tracking Post-Remate

### `filtrador_saldos.py` (~3500 líneas, creado 2026-03-26)

**Arquitectura:**
- Sistema 100% independiente de M1-M5
- 5 workers paralelos via subprocess (cada uno con su propio Playwright/Chromium)
- Distribución fija por Corte de Apelaciones
- Lee reportes M5 de `Reportes/` como entrada

**Estado de componentes:**
1. ✅ **Merge (Paso 0):** Importa causas de reportes M5, dedup ROL+AÑO, excluye RM/DESCONOCIDA
2. ✅ **Filtro 1 (4 señales):** Lee tabla `#historiaCiv` cuaderno Apremio sin PDFs
3. ✅ **Filtro 2 (acta de remate):** Detección en tabla + descarga PDF + OCR + cargo al crédito + monto
4. ✅ **Filtro 3 (liquidación):** Descarga PDF + regex Tipo A/B + sanidad

### Filtro 1: Señales en tabla

| Señal | Patrón | Decisión |
|---|---|---|
| 1 | "no postores" | ELIMINAR |
| 2 | Resolución + "Suspende/Reprograma" | ELIMINAR |
| 3 | "Nuevo día y hora" | ELIMINAR |
| 4 | Keywords liquidación (12 variantes) | PENDIENTE_LIQUIDACION |

Keywords liquidación: "ordena liquidar el crédito", "liquidacion (credito)", "pone en conocimiento liquidación de crédito", "solicita liquidación", y variantes sin tildes.

### Filtro 2: Acta de remate (PDF)

**Arquitectura de dos pasadas en `_evaluar_filtros_tabla()`:**
1. Pasada 1: buscar "acta de remate" en CUALQUIER fila → si existe, SIEMPRE retornar NECESITA_PDF_ACTA
2. Pasada 2: solo si NO hay acta → evaluar señales F1 normales

Esto garantiza que "cargo al crédito" SIEMPRE se verifica, incluso si hay liquidación más reciente.

**Descarga de PDF:** Form con `action="docuS.php"` + hidden JWT token → nueva pestaña → captura con Playwright

**Análisis del PDF (`_analizar_pdf_acta()`):**
- PyMuPDF extrae texto nativo
- Si no hay contenido de acta (no `$` ni `adjudica`) → OCR fallback con pytesseract (DPI 150, timeout 30s/página)
- Detección cargo al crédito: regex `r'cargo\s+(?:a\s+(?:su|sus|los?)|al)\s+cr[eé]ditos?'` IGNORECASE
- Extracción monto: regex CLP `$XX.XXX.XXX`, busca post-"adjudica", primer monto después
- Tope de sanidad: MONTO_MAXIMO_RAZONABLE = 2_000_000_000

**Lógica de decisión:**
- cargo al crédito → ELIMINAR (sin excedente posible)
- monto < deuda → ELIMINAR
- monto >= deuda → PENDIENTE_LIQUIDACION (posible excedente)

### Filtro 3: Liquidación (PDF)

**Selector:** TODAS las causas con estado PENDIENTE_LIQUIDACION.

**Búsqueda en tabla OJV:**
Solo busca filas con tramite "Liquidacion (Credito)" o "Liquidación (Crédito)".
NO buscar "Ordena liquidar" ni "Pone en conocimiento" (esas son resoluciones
del juez, no el documento con cálculos).

Si la fila no existe → causa se queda en PENDIENTE_LIQUIDACION (el departamento
de liquidaciones aún no ha subido el PDF). Se reintenta en el próximo run.

**Descarga PDF:** Misma mecánica que F2 (form docuS.php + JWT).

**Dos tipos de liquidación:**

| Tipo | Contenido PDF | Regex | Decisión |
|------|--------------|-------|----------|
| A | Incluye precio de remate + cálculo de saldo | `Saldo a favor del (Ejecutado\|Demandado)` | saldo > 0 → EXCEDENTE_CONFIRMADO |
| B | Solo calcula deuda total (sin comparar remate) | `(Crédito\|Capital) adeudado al` | delta = acta - crédito → según resultado |

**Regex Tipo A (re.DOTALL):**
```python
r'[Ss]aldo\s+a\s+favor\s+del\s+(?:[Ee]jecutado|[Dd]emandado).*?(\d{1,3}(?:\.\d{3})+)'
```

**Regex Tipo B:** Busca monto inmediatamente después de "Crédito adeudado al" o
"CAPITAL ADEUDADO AL". Si hay múltiples matches, toma el mayor (es el total).

**Lógica de decisión Tipo B:**
- Si tiene monto_acta_remate: delta = acta - crédito_adeudado
  - delta > 0 → EXCEDENTE_CONFIRMADO
  - delta <= 0 → ELIMINAR
- Si NO tiene monto_acta → PENDIENTE_REVISION_MANUAL

**Regla de sanidad Tipo B:**
Si crédito_adeudado < deuda_original × 0.5 → error de extracción → REVISION_MANUAL.
(Imposible que la deuda con intereses+costas sea menor que la mitad de la deuda original)

**OCR fallback para liquidaciones:** Criterio de contenido útil:
"saldo", "liquidaci", "capital adeudado", "interés/interes", "costas", "$"+dígitos.

**Learning clave:** "Ordena liquidar el crédito" ≠ "Liquidacion (Credito)".
El primero es la orden del juez, el segundo es el documento con los cálculos.
F1 detecta ambos como señal, pero F3 solo descarga el segundo.

### Estados del Excel madre
- `PENDIENTE_FILTRO1` — recién importada, sin procesar
- `PENDIENTE_ACTA` — pasó F1, sin acta ni liquidación en tabla
- `NECESITA_PDF_ACTA` — acta detectada en tabla, PDF pendiente
- `PENDIENTE_LIQUIDACION` — posible excedente, esperando F3
- `PENDIENTE_REVISION_MANUAL` — liquidación concursal Ley 20.720
- `EXCEDENTE_CONFIRMADO` — saldo > 0, liquidación procesada y verificada
- `ELIMINAR` — sin excedente (cargo al crédito, monto < deuda, suspensión, no postores)

### Excel madre: pestañas
- `_datos_internos` — todas las causas con historial completo
- `Causas con Saldo` — solo estados activos (excluye ELIMINAR y REVISION_MANUAL)
- `Por Antigüedad` — ordenada por días desde remate (priorización)
- `Revisión Manual` — causas que requieren intervención humana

### Flags de ejecución
```
--primera-run --solo-merge          # solo importar reportes
--primera-run --solo-filtro1        # solo F1 sin PDFs
--reaudit --workers 5               # reprocesar PENDIENTE_ACTA + PENDIENTE_LIQUIDACION sin monto
--recheck-cargo                     # re-analizar PDFs locales con regex actualizado (sin OJV)
--solo-filtro3                      # solo F3 sin F1/F2
--recheck-liq                       # re-analizar PDFs liquidación locales (sin OJV)
--skip-pdf                          # debugging: no descargar PDFs
--workers N                         # número de workers paralelos
```

### Resultados acumulados (2026-03-31, post F1+F2+F3 + eliminación lote)

| Métrica | Valor |
|---|---|
| Causas en Excel madre | 168 |
| EXCEDENTE_CONFIRMADO | 2 (~$84.4M total) |
| PENDIENTE_LIQUIDACION | 16 (esperando PDF en tabla OJV) |
| PENDIENTE_REVISION_MANUAL | 2 (1 Tipo B sin acta + 1 concursal) |
| PENDIENTE_ACTA | 6 |
| PENDIENTE_FILTRO1 | 129 |
| ELIMINAR | 20 (6 F1+F2+F3 + 8 revisión abogados + 6 verificación manual) |
| PDFs actas descargados | ~30 |
| PDFs liquidaciones descargados | ~10 |

### Excedentes confirmados
| Causa | Tipo | Excedente | Detalle |
|---|---|---|---|
| C-1529-2023 | A | $61,986,279 | Saldo a favor del ejecutado |
| C-2540-2025 | B | $22,385,979 | acta $61.2M - crédito $38.8M |

### Causas eliminadas por revisión (fix_eliminar_lote.py)
**8 revisión abogados** (sin futuro comercial): C-3323-2024, C-7205-2024, C-1810-2025,
C-5950-2020, C-2601-2023, C-1833-2025, C-4885-2023, C-2431-2024.

**6 verificación manual:**
| Causa | Razón |
|---|---|
| C-7366-2024 | CCC folio 77. Banco Itaú $321M cargo a sus créditos |
| C-93-2024 | CCC folio 75. FactorOne $215M cargo a su crédito. Saldo $4.8M no perseguible |
| C-3535-2025 | CCC folio 21. Banco Santander $100M cargo a su crédito |
| C-4836-2022 | CCC folio 233. Lote C $30.7M + Lote D $36.8M ambos cargo al crédito |
| C-867-2024 | Suspensión de remate folio 220. Sin acta |
| C-1138-2024 | Acta $110M < crédito adeudado $166M. Sin excedente |

### Bugs resueltos durante F2
- ✅ Modal `#modalDetalleCivil` — 3 niveles de fallback
- ✅ Cuaderno Apremio — selección explícita
- ✅ Workers subprocess — PIDs independientes
- ✅ Cargo al crédito plural/variantes — regex flexible
- ✅ OCR no se activaba (footer PJUD daba texto) — check por contenido, no longitud
- ✅ OCR crash/timeout — try/except por página, DPI 150, timeout 30s
- ✅ MuPDF spam — `fitz.TOOLS.mupdf_display_errors(False)`
- ✅ Monto OCR irrazonable ($10.1B) — tope MONTO_MAXIMO_RAZONABLE = 2B
- ✅ PermissionError Excel abierto — `_guardar_excel_con_retry()` 3 intentos + backup
- ✅ TypeError int→str columna — `str(nuevo_monto)` en recheck-cargo
- ✅ Reaudit scope — incluye PENDIENTE_LIQUIDACION sin monto_acta
- ✅ Señal 4 oculta acta — dos pasadas en `_evaluar_filtros_tabla()` (acta primero)

### Bugs resueltos durante F3
- ✅ Argumentos corte/tribunal invertidos en buscar_causa() — F3 pasaba tribunal donde iba corte
- ✅ Selector restrictivo — solo veía 12 de 29 causas, perdía $742M potencial
- ✅ PDF equivocado — descargaba "Ordena liquidar" en vez de "Liquidacion (Credito)"
- ✅ Regex "ejecutado" no matcheaba "demandado" — C-93 usa "SALDO A FAVOR DEL DEMANDADO"
- ✅ Regex Tipo B "último monto grande" capturaba costas — cambiado a monto post-"crédito adeudado"
- ✅ Excedente falso C-93 ($210M→$4.8M) — regex no matcheaba "demandado", Tipo B agarró costas
- ✅ Excedente falso C-2609 ($178M→ELIMINAR) — Tipo B agarró costas $1.1M en vez de crédito $417M
- ✅ C-2470 delta negativo — acta $51M < deuda $68M, movida a ELIMINAR
- ✅ 10 causas atrapadas en REVISION_MANUAL por run buggy — revertidas con fix_revert_f3.py
- ✅ PDFs basura del run buggy en carpeta liquidaciones/ — borrados manualmente

### Learnings clave F3
- "Ordena liquidar" ≠ "Liquidacion (Credito)" — el primero es orden, el segundo es documento
- "Ejecutado" y "Demandado" son sinónimos — ambos aparecen en liquidaciones reales
- Liquidaciones de Santiago (Unidad de Liquidaciones) suelen ser Tipo A (saldo directo)
- Liquidaciones de regiones suelen ser Tipo B (solo crédito adeudado)
- Regla de sanidad (crédito < 50% deuda = error) habría atrapado ambos excedentes falsos
- PDFs de liquidación real: 300-500KB, 3000-6000 chars. Resoluciones: 65-100KB, 700-1000 chars

### Learnings clave F2
- "Cargo al crédito" tiene muchas variantes: "su/sus/los/al" + "crédito/créditos" → regex flexible
- PDFs del PJUD: ~60% texto nativo, ~40% escaneados como imagen
- Footer de firma PJUD tiene texto nativo en PDFs escaneados → no usar longitud para detectar
- `_evaluar_filtros_tabla()` debe buscar acta PRIMERO en pasada separada (cargo prevalece sobre todo)
- `--recheck-cargo` útil cuando se actualiza el regex sin re-descargar PDFs

---

## Módulo 4 — ⛔ PERMANENTEMENTE ABANDONADO

**NO usar. NO reactivar. NO sugerir alternativas de tasación automatizada.**

---

## Costos API

- **Pipeline v2 run semanal:** ~$0.15 USD (~64 llamadas Haiku)
- **Filtrador de Saldos:** $0 (solo Playwright + scraping + OCR local)

---

## Dependencias

```
pymupdf (fitz), pandas, openpyxl, playwright, numpy, requests, rapidfuzz, anthropic, python-docx, pytesseract, Pillow
```

Requiere Tesseract-OCR instalado en Windows con idioma español (spa).
