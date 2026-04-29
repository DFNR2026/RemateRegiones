# PROGRESO — Sistema de Análisis de Remates Judiciales Chile

Última actualización: 2026-04-11 (M2 workers + WAF bypass + Filtrador reaudit)

---

## Archivos del proyecto

### `config.py`
Configuración global: rutas, claves API, constantes, CAUSAS_IGNORADAS.

### `modulo1_parser.py` ✅ (v1 — respaldo)
### `v2_experimental/modulo1_v2.py` ✅ EN PRODUCCIÓN (Regex + Haiku)
### `modulo2_ojv.py` ✅ COMPLETADO — 5 workers subprocess round-robin
### `modulo3_extractor.py` ✅ COMPLETADO
### `modulo5_reporte.py` ✅ COMPLETADO
### `main.py` ✅ COMPLETADO — llama M2 con n_workers=5
### `ojv_remates.py` (v10.0+) — MOTOR OJV BASE — NO reescribir
### `causas_ojv.xlsx` — BD interna (235 tribunales + historial)
### `test_m2_sandbox.py` — test M2 sin correr M1 (sin costo API)
### `Detector_Excedentes.bat` — ejecución semanal filtrador
### `Regiones_docxToExcel.bat` — ejecución semanal pipeline M1-M5

---

## Workers M2 (agregado 2026-04-09)

- 5 workers subprocess, distribución round-robin
- `launch_persistent_context` + `channel="chrome"` + `playwright-stealth`
- Perfiles `.chrome-profile-wN` limpiados automáticamente al inicio
- Stagger 3s entre lanzamientos
- Resultados: 67 causas en ~3 min (vs ~30 min secuencial)

---

## WAF Bypass (incidente 2026-04-09/10)

PDJ activó WAF F5 BIG-IP que bloqueaba Chromium de Playwright.
Solución: `launch_persistent_context` + `channel="chrome"` + `playwright-stealth`.
`navegar_a_consulta()` reescrita: `home/index.php` → click "Consulta causas" → select "Civil".
Ver `INCIDENTE_WAF_2026-04-09.md` para detalle completo.

---

## Filtrador de Saldos

### `filtrador_saldos.py` (~4500 líneas)

**Estado de componentes:**
1. ✅ **Merge (Paso 0):** Importa causas de reportes M5
2. ✅ **Filtro 1 (4 señales):** Lee tabla `#historiaCiv` cuaderno Apremio
3. ✅ **Filtro 2 (acta de remate):** Detección + PDF + OCR + cargo al crédito
4. ✅ **Filtro 3 (liquidación):** PDF + regex Tipo A/B + sanidad

### Resultados acumulados (2026-04-10, post reaudit)

| Métrica | Valor |
|---|---|
| Causas en Excel madre | ~270 |
| EXCEDENTE_CONFIRMADO | 2 (~$84.3M) |
| PENDIENTE_LIQUIDACION | 14 |
| PENDIENTE_REVISION_MANUAL | 6 |
| PENDIENTE_ACTA | 6 |
| PENDIENTE_FILTRO1 | 169 |
| ELIMINAR | 50 |

### Excedentes confirmados
| Causa | Tipo | Excedente |
|---|---|---|
| C-1529-2023 | A | $61,986,279 |
| C-2540-2025 | B | $22,385,979 |

---

## Módulo 4 — ⛔ PERMANENTEMENTE ABANDONADO

---

## Costos API

- **Pipeline v2 run semanal:** ~$0.15 USD (~64 llamadas Haiku)
- **Filtrador de Saldos:** $0 (solo Playwright + scraping + OCR local)

---

## Dependencias

```
pymupdf (fitz), pandas, openpyxl, playwright, numpy, requests, rapidfuzz,
anthropic, python-docx, pytesseract, Pillow, playwright-stealth
```

Requiere Tesseract-OCR (Windows, idioma español).
Requiere Google Chrome instalado (`channel="chrome"`).
