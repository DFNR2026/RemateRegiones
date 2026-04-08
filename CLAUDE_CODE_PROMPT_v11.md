# PROYECTO: Sistema Automatizado de Análisis de Remates Judiciales - Chile
# Versión 11 del prompt maestro (actualizado 2026-03-31)

## CONTEXTO DEL NEGOCIO

Inversionista inmobiliario chileno que analiza remates judiciales para identificar excedentes (cuando la subasta supera la deuda). Proceso ~200 causas semanales del Diario P&L. Foco: propiedades fuera de la Región Metropolitana.

---

## ARQUITECTURA: DOS SISTEMAS INDEPENDIENTES

### Sistema 1: Pipeline Principal (M1-M5) ✅ EN PRODUCCIÓN

```
D:\Remates\
├── main.py                ← orquestador M1→M2→M3→M5
├── v2_experimental/modulo1_v2.py  ← parser v2 EN PRODUCCIÓN (Regex + Haiku)
├── modulo2_ojv.py         ← consulta OJV via Playwright
├── modulo3_extractor.py   ← extrae montos de deuda desde PDFs
├── modulo5_reporte.py     ← reporte final Excel
├── ojv_remates.py         ← motor OJV base (v10.0) — NO REESCRIBIR
├── causas_ojv.xlsx        ← BD INTERNA (235 tribunales + historial)
├── config.py              ← claves API, rutas, constantes
└── Reportes/              ← salida M5 = entrada del Filtrador
```

### Sistema 2: Filtrador de Saldos ✅ F1+F2+F3 FUNCIONALES

```
D:\Remates\
├── filtrador_saldos.py    ← ~3500 líneas, 5 workers subprocess
├── Detector_Excedentes.bat   ← ejecución semanal (merge+F1+F2+F3)
├── temp_workers/          ← JSONs temporales workers
├── audit_html/            ← outerHTML tabla Apremio
├── Descargas/actas/       ← PDFs de actas descargados por F2
├── logs liqui/            ← logs del filtrador
└── Causas con liq/
    ├── Causas_posible_saldo.xlsx   ← Excel madre (168 causas)
    ├── Causas_con_liquidacion.xlsx ← excedentes confirmados + rev manual
    ├── liquidaciones/              ← PDFs de "Liquidacion (Credito)"
    └── liquidaciones_raw/          ← texto extraído de cada PDF
```

### ⛔ MÓDULO 4 — PERMANENTEMENTE ABANDONADO

---

## FILTRADOR DE SALDOS — DOCUMENTACIÓN COMPLETA

### Objetivo
Buscar excedentes post-remate. Documento clave: la **Liquidación** (~14 días post-remate).

### Vocabulario judicial
- Ejecutado = Demandado | Ejecutante = Demandante (banco)
- Acta de Remate = monto de adjudicación | Liquidación = saldo final
- Cargo al crédito = banco se adjudicó sin plata real → sin excedente → ELIMINAR

### Flujo semanal
```
1. DOCX semanal → D:\Remates\Diarios\
2. Doble-click Regiones_docxToExcel.bat → pipeline M1-M5 → Reportes\
3. Doble-click Detector_Excedentes.bat → F1+F2+F3 → Causas_con_liquidacion.xlsx
```

### Ejecución
```
python filtrador_saldos.py --workers 5                    # run normal
python filtrador_saldos.py --reaudit --workers 5          # reprocesar pendientes
python filtrador_saldos.py --recheck-cargo                # re-analizar PDFs locales
python filtrador_saldos.py --skip-pdf --workers 5         # sin descarga PDFs
python filtrador_saldos.py --solo-filtro3 --workers 5     # solo F3
python filtrador_saldos.py --recheck-liq                  # re-analizar liquidaciones locales
Detector_Excedentes.bat                                      # doble-click: merge+F1+F2+F3
```

### Paralelización: 5 workers por Corte (subprocess)

```python
WORKER_CORTES = {
    1: ["C.A. de La Serena"],
    2: ["C.A. de Valparaíso", "C.A. de Punta Arenas", "C.A. de Talca"],
    3: ["C.A. de Concepción", "C.A. de Coyhaique", "C.A. de Chillán"],
    4: ["C.A. de Antofagasta", "C.A. de Rancagua"],
    5: ["C.A. de Iquique", "C.A. de Temuco", "C.A. de Copiapó", "C.A. de Puerto Montt"],
}
```

### Función clave: `_evaluar_filtros_tabla()` — DOS PASADAS

```
Pasada 1: Buscar "acta de remate" en CUALQUIER fila
          → Si existe: retornar NECESITA_PDF_ACTA (cargo al crédito prevalece sobre todo)

Pasada 2: Solo si NO hay acta → evaluar señales F1 (más reciente primero):
          → Liquidación concursal → PENDIENTE_REVISION_MANUAL
          → Keywords liquidación (12 variantes) → PENDIENTE_LIQUIDACION
          → "no postores" → ELIMINAR
          → Resolución + "Suspende/Reprograma" → ELIMINAR
          → "Nuevo día y hora" → ELIMINAR
```

### Filtro 2: Análisis PDF del acta (`_analizar_pdf_acta()`)

1. PyMuPDF extrae texto nativo
2. Si no hay `$` ni `adjudica` en texto → OCR fallback (pytesseract, DPI 150, timeout 30s/pág)
3. Detección cargo: `r'cargo\s+(?:a\s+(?:su|sus|los?)|al)\s+cr[eé]ditos?'` IGNORECASE
4. Extracción monto: regex `$XX.XXX.XXX`, post-"adjudica", tope 2B

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

### Descarga PDF desde tabla OJV

Cada fila de `#historiaCiv` tiene un `<form action="docuS.php">` con `<input hidden name="dtaDoc" value="JWT">`.
Click en el ícono PDF submite el form → nueva pestaña con el PDF.

### Estados del Excel madre
- `PENDIENTE_FILTRO1` — sin procesar
- `PENDIENTE_ACTA` — sin acta ni liquidación
- `NECESITA_PDF_ACTA` — acta detectada, PDF pendiente
- `PENDIENTE_LIQUIDACION` — posible excedente, esperando F3
- `PENDIENTE_REVISION_MANUAL` — liquidación concursal
- `EXCEDENTE_CONFIRMADO` — saldo > 0, liquidación procesada y verificada
- `ELIMINAR` — sin excedente (cargo, monto<deuda, suspensión, no postores)

### Excel madre: pestañas
- `_datos_internos` — todas las causas
- `Causas con Saldo` — solo estados activos (excluye ELIMINAR/REVISION_MANUAL)
- `Por Antigüedad` — ordenada por días desde remate
- `Revisión Manual` — requiere intervención humana

### Señales informativas (no cambian estado, solo detalle)
- "giro cheque" / "gírese" / "consigna precio de remate"
- "objeta liquidación" / "nulidad de lo obrado"
- "da cuenta de pago" / "abandono procedimiento" / "desistimiento"

---

## RESULTADOS ACUMULADOS (2026-03-31, post eliminación lote)

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

---

## INSTRUCCIONES PARA CLAUDE CODE

### Reglas generales
1. **NO reescribir ojv_remates.py desde cero.**
2. **NO ejecutar main.py, filtrador_saldos.py, ni ningún módulo.** Diego ejecuta manualmente.
3. **Montos SIEMPRE en CLP.** Convertir UF via mindicador.cl.
4. **M4 NO EXISTE.**
5. **config.py** contiene TODAS las claves, rutas y constantes.
6. **Encoding logs: solo ASCII** en f-strings. Windows cp1252.
7. **"Ejecutado" = "Demandado"** — sinónimos.

### Reglas del Filtrador
8. **Sistema 100% independiente** de M1-M5.
9. **5 workers via subprocess**, distribución fija por Corte.
10. **SIEMPRE cerrar modal** `#modalDetalleCivil` después de leer tabla.
11. **SIEMPRE seleccionar cuaderno Apremio** en `#selCuaderno`.
12. **Cargo al crédito = ELIMINAR siempre** — regex flexible, no strings exactos.
13. **MONTO_MAXIMO_RAZONABLE = 2_000_000_000** — descartar montos OCR irrazonables.
14. **`_evaluar_filtros_tabla()` SIEMPRE busca acta primero** (dos pasadas).
15. **OCR fallback**: detectar por contenido (`$` + `adjudica`), no por longitud de texto.
16. **Suprimir MuPDF warnings**: `fitz.TOOLS.mupdf_display_errors(False)`
17. **PermissionError Excel**: `_guardar_excel_con_retry()` con 3 intentos + backup.
18. **Playwright sync API NO soporta multithreading** — solo subprocess/multiprocessing.
19. **F3 busca SOLO "Liquidacion (Credito)"** en tabla OJV — NO "Ordena liquidar"
20. **"Ejecutado" = "Demandado"** en regex de liquidación — ambos son válidos
21. **Tipo B: monto post-"crédito adeudado"** — NUNCA usar "último monto grande"
22. **Sanidad Tipo B: crédito < 50% deuda = error** → REVISION_MANUAL
23. **Fila no encontrada en F3 = PENDIENTE** (no REVISION_MANUAL) — el PDF puede no existir aún
24. **Excel Causas_con_liquidacion.xlsx con separadores por fecha** de ejecución

---

## DEPENDENCIAS

```
pymupdf (fitz), pandas, openpyxl, playwright, numpy, requests, rapidfuzz, anthropic, python-docx, pytesseract, Pillow
```

Requiere Tesseract-OCR instalado en Windows con idioma español (spa).
