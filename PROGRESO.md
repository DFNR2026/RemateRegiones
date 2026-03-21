# PROGRESO — Sistema de Análisis de Remates Judiciales Chile

Última actualización: 2026-03-20 (v2 Regex+Haiku default, 15 fixes, Colina en REFERENCIA)

---

## Archivos del proyecto

### `config.py`
Configuración global del proyecto: rutas, claves API, constantes.
- `DIARIOS_DIR`, `DESCARGAS_DIR`, `CAUSAS_XLSX`
- `ANTHROPIC_API_KEY` (Claude API para M1)
- `DEMANDANTES_EXCLUIDOS`, `CORTES_RM`
- `CAUSAS_IGNORADAS` — blacklist de causas con cuadernos restringidos/inaccesibles en OJV

### `modulo1_parser.py` ✅ COMPLETADO (v1 — Sonnet, fallback con `--v1`)
Parser de avisos de remate desde PDFs diarios O DOCX semanal consolidado.
Usa Claude Sonnet 4.6 para extracción de todos los campos.

**Output:** Lista de dicts con campos: `rol, año, corte, tribunal, demandante, demandado, direccion, comuna, region_rm, fecha_publicacion`

### `v2_experimental/modulo1_v2.py` ✅ DEFAULT (v2 — Regex+Haiku)
Parser optimizado que reemplaza a v1 como default para `--docx`.
- **Estrategia Regex-Flow + Haiku**: ROL/tribunal/demandante con regex puro, dirección/comuna con Haiku solo cuando regex falla
- **Costo**: ~$0.05/run vs ~$3/run (v1) — ahorro ~60x
- **Velocidad**: ~3-4x más rápido que v1
- **Match rate vs v1**: 85.7% (diferencias son en detalle de dirección, no en campos críticos)

**15 fixes aplicados (Fix 1-15, Fix 17-18):**
- Fix 1: Corte DESCONOCIDA por falta de buscar_corte
- Fix 2: Dirección regex captura basura (validación longitud + dígito)
- Fix 3: Demandante regex no captura bancos (BCI, Scotiabank, etc.)
- Fix 4: Aplanar texto (\n→espacio) antes de regex
- Fix 5: Validación dirección con números en palabras ("número", "lote", etc.)
- Fix 6: Tribunal captura texto de más (truncar en keywords remate)
- Fix 7: Logging propio a v2_experimental/logs/
- Fix 8: Comuna con artículo ("La Serena" no "Serena")
- Fix 9: Demandante slash/unquoted caratulada
- Fix 10: Haiku retorna dirección + comuna
- Fix 11: Comuna exceso captura ("Bulnes sector..." → "Bulnes")
- Fix 12: Comuna vacía → extraer de tribunal normalizado (REFERENCIA)
- Fix 13: Haiku fallback para tribunal cuando regex v1+v2 fallan
- Fix 14: Ordinal "01°" → "1°" normalización
- Fix 15: Juzgado de Letras de Colina agregado a REFERENCIA
- Fix 17: fecha_remate revertido (no se usa en pipeline)
- Fix 18: Logging tribunal regex v1/v2/Haiku para diagnóstico

**Resultados run producción DOCX (semana 16-20 marzo 2026):**
- 184 bloques procesados → 63 causas nuevas
- 58/63 documentos descargados de OJV (92%)
- 234 tribunales en REFERENCIA (incluye 29°/30° Santiago + Colina)
- 1 sola causa con corte DESCONOCIDA (C-240-2025 Linares)

**Fixes históricos v1:**
- Encoding fix: `→` → `->` para Windows cp1252
- Ordinal/City recovery en buscar_corte()
- Ordinales textuales hasta "Trigésimo"
- RapidFuzz reemplazó difflib SequenceMatcher

### `modulo2_ojv.py` ✅ COMPLETADO
Wrapper del motor OJV para el pipeline.

**Qué hace:**
- Recibe lista de dicts del Módulo 1 (no lee Excel)
- Expone `procesar_causas_ojv(causas) -> list[dict]`
- Enriquece cada causa con: `tipo_procedimiento`, `tipo_documento`, `descargado`, `ruta_pdf`
- **Extracción de litigantes DTE/DDO desde OJV**: navega a pestaña `#litigantesCiv`, extrae nombre completo
- **Filtro blacklist**: causas en `CAUSAS_IGNORADAS` se saltan
- **Filtro procedimientos descartados**: liquidación simplificada/concursal, ordinario, partición, arbitral, arrendamiento, monitorio

### `modulo3_extractor.py` ✅ COMPLETADO
Extrae montos de deuda de los PDFs descargados por el Módulo 2.

**Qué hace:**
1. Para cada causa con `descargado=True`, abre el PDF con PyMuPDF
2. Según `tipo_documento`: busca patrones de monto en mandamiento o bases de remate
3. Extrae monto en UF o CLP
4. UF → CLP usando API `mindicador.cl/api/uf` (con caché + fallback $38.500)

**Output:** enriquece con `monto_deuda_clp` (int), `monto_original` (str)

### `modulo5_reporte.py` ✅ COMPLETADO
- **`actualizar_historial(causas)`**: APPEND a hoja CAUSAS, deduplicación por (ROL, AÑO)
- **`generar_reporte(causas)`**: Excel con 3 pestañas (Resumen, RM, Regiones), formato condicional por ratio

### `main.py` ✅ COMPLETADO
Orquestador M1 → M2 → M3 → M5.

**Uso:**
```
python main.py --docx "ruta.docx"             # DOCX semanal (v2 Regex+Haiku por defecto)
python main.py --v1 --docx "ruta.docx"        # DOCX forzando Sonnet (v1)
python main.py --docx "ruta.docx" --limpiar-historial  # test sin afectar producción
python main.py                                  # PDFs diarios (fallback, usa v1 Sonnet)
python main.py --sin-ojv                        # omite M2
python main.py --hasta N                        # detiene tras Módulo N
python main.py --silencio                       # solo resúmenes
```

### `ojv_remates.py` (v10.0+) — MOTOR OJV BASE
**NO reescribir desde cero.** Automatización OJV con Playwright + RapidFuzz.

### `causas_ojv.xlsx`
BD interna con dos hojas:
- **REFERENCIA**: 234 filas con mapeo tribunal → corte (actualizada 2026-03-20: +29°/30° Santiago, +Colina)
- **CAUSAS**: historial ROLes procesados

### `limpiar_cache.py`
Limpieza de caché antes de test runs.

---

## Archivos eliminados (limpieza 2026-03-14)

| Archivo | Motivo |
|---|---|
| `modulo1_mercurio.py` | Parser alternativo experimental, nunca en producción |
| `analizar_log.py` | Script diagnóstico puntual |
| `causas_antes.json`, `causas_despues.json` | Artefactos de debug |
| `tasacion_cache.db` (74MB) | Caché de M4, permanentemente abandonado |
| `Reporte_2026-02-28.xlsx`, `Reporte_2026-03-01.xlsx` | Reportes viejos en raíz (movidos a Reportes/) |
| `TEST_RESUMEN_16_CAUSAS.docx`, `TEST_SANDBOX_24_CAUSAS.docx` | Archivos de test sandbox |
| `backup_20260301/` | Backup viejo pre-refactor |
| `la API/` | Carpeta vacía experimental |
| `htmls de PDJ/` | HTMLs de referencia del Poder Judicial |
| `.tessdata/` | Datos OCR (innecesario con DOCX) |
| `Diarios test/` | PDFs de prueba |
| `Diarios_Procesados/` | PDFs ya procesados |
| `__pycache__/` | Caché Python |

---

## Módulo 4 — ⛔ PERMANENTEMENTE ABANDONADO

**NO usar. NO reactivar. NO sugerir alternativas de tasación automatizada.**

---

## Costos API

- **Run DOCX semanal v2 (184 bloques):** ~$0.05 USD en tokens Haiku (solo ~50% de causas necesitan Haiku)
- **Run DOCX semanal v1 (184 bloques):** ~$3 USD en tokens Sonnet 4.6 (fallback con `--v1`)
- **Tiempo M1 v2:** ~30s (regex puro + ~30 llamadas Haiku)
- **Tiempo M1 v1:** ~5 min (184 llamadas API secuenciales)
- **Tiempo M2:** ~21 min (51-58 causas en OJV)
- **Tiempo total pipeline v2:** ~22 min

---

## Dependencias

```
pymupdf (fitz), pandas, openpyxl, playwright, numpy, requests, rapidfuzz, anthropic, python-docx
```

## Notas técnicas clave

- **DOCX semanal es la fuente principal** desde 2026-03-14. PDFs diarios quedan como fallback.
- **Flujo de datos en memoria**: lista de dicts entre módulos, sin archivos intermedios
- **UF**: en PDFs aparece como número con coma decimal (ej: "1.767,802476 UF")
- **~7.3% tasa de no-descarga OJV** (4/55): 2 procedimientos no aplicables, 1 cuaderno restringido, 1 tribunal no reconocido
- **Montos en CLP siempre**: UF se convierte via mindicador.cl
- **`--limpiar-historial`** para tests sin contaminar producción
