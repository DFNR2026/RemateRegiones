# PROGRESO — Sistema de Análisis de Remates Judiciales Chile

Última actualización: 2026-03-14 (migración DOCX semanal, ordinal recovery, city recovery, 29°-30° Santiago)

---

## Archivos del proyecto

### `config.py`
Configuración global del proyecto: rutas, claves API, constantes.
- `DIARIOS_DIR`, `DESCARGAS_DIR`, `CAUSAS_XLSX`
- `ANTHROPIC_API_KEY` (Claude API para M1)
- `DEMANDANTES_EXCLUIDOS`, `CORTES_RM`
- `CAUSAS_IGNORADAS` — blacklist de causas con cuadernos restringidos/inaccesibles en OJV

### `modulo1_parser.py` ✅ COMPLETADO
Parser de avisos de remate desde PDFs diarios O DOCX semanal consolidado.

**Qué hace:**
1. Lee hoja REFERENCIA de `causas_ojv.xlsx` (233 tribunales → cortes)
2. Lee hoja CAUSAS (historial de ROLes procesados)
3. **Modo DOCX (principal desde 2026-03-14):** Lee `.docx` semanal con `python-docx`, separa por encabezados de fecha (`16 MARZO`, `17 MARZO`, etc.), cada párrafo no-vacío es una causa completa
4. **Modo PDF (fallback):** Para cada PDF en `Diarios/`: extrae texto con PyMuPDF, separa en bloques por marcadores
5. Filtra bloques de jueces árbitros/partidores (masculino Y femenino, tolerante a errores OCR)
6. Por cada bloque: envía a Claude API (`claude-sonnet-4-6`) para extraer ROL, tribunal, demandante, dirección y comuna
7. **Post-procesamiento `_limpiar_tribunal()`**: une guiones silábicos, elimina direcciones físicas del nombre del tribunal, normaliza capitalización
8. **Filtro post-extracción**: si Claude devuelve tribunal tipo "Jueza Partidora...", "Juez Árbitro...", la causa se descarta
9. Deduplica por ROL entre bloques de la semana y contra historial
10. Filtra: Banco Estado, causas pre-2018
11. Mapea tribunal → corte con **RapidFuzz** (`fuzz.token_set_ratio`, umbral 80.0)
12. **Validación ordinal post-matching**: si el número ordinal no coincide, rechaza el match e **intenta recovery** filtrando REFERENCIA por tribunales con el ordinal correcto
13. **`_extraer_ordinal()`**: reconoce ordinales numéricos (1°, 2º, 29°) Y textuales ("Primer", "Séptimo", "Vigésimo Noveno", etc.) via diccionario texto→número
14. **Validación de ciudad post-matching**: compara ciudades entre tribunal del PDF y tribunal candidato (con normalización de tildes). Si difieren, penaliza 0.7x. Si cae bajo umbral, **intenta city recovery** filtrando REFERENCIA por tribunales que contengan la ciudad del PDF
15. **Retry dirección**: si Claude no extrajo dirección pero sí ROL, hace un segundo intento con prompt focalizado
16. Determina region_rm según corte (C.A. de Santiago o C.A. de San Miguel)
17. **Campo `fecha_publicacion`**: extraído del encabezado de fecha del DOCX (ej: "16 MARZO 2026")

**Output:** Lista de dicts con campos: `rol, año, corte, tribunal, demandante, demandado, direccion, comuna, region_rm, fecha_publicacion`

**Resultados run producción DOCX (semana 16-20 marzo 2026):**
- 184 bloques procesados → 55 causas nuevas
- 117 sin ROL (bloques de ruido: encabezados, publicidad, párrafos vacíos)
- 10 filtradas Banco Estado, 2 filtradas pre-2018
- 51/55 documentos descargados de OJV (92.7%)
- 51/51 montos extraídos (100%)
- 1 sola causa con corte DESCONOCIDA (C-240-2025 Linares, edge case de score 79.2%)

**Fixes aplicados (2026-03-14, migración DOCX):**
- **Nueva función `parsear_docx_semanal()`**: lee DOCX con `python-docx`, convive con `parsear_diarios()` como fallback
- **Encoding fix**: reemplazado `→` (U+2192) por `->` en logs para compatibilidad con consola Windows cp1252
- **Ordinal recovery**: cuando ORDINAL MISMATCH rechaza el top match, re-busca en subset de REFERENCIA filtrado por ordinal correcto
- **Ordinales textuales**: diccionario "Primer"→1, "Séptimo"/"Sétimo"→7, "Vigésimo Noveno"→29, etc.
- **City recovery**: cuando CITY MISMATCH rechaza, re-busca filtrando REFERENCIA por ciudad del PDF
- **Normalización tildes en ciudades**: `unicodedata.normalize()` para comparar "quilpué" == "quilpue"
- **Tribunales agregados a REFERENCIA**: 29° y 30° Juzgado Civil de Santiago (faltaban, lista iba hasta 28°)

**Fixes históricos:**
- Upgrade Haiku → Sonnet 4.6 (eliminó alucinaciones graves)
- Claude API reemplazó regex frágil para extracción de campos
- RapidFuzz reemplazó difflib SequenceMatcher
- Filtro jueces partidores ampliado a femenino + tolerancia OCR

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
python main.py --docx "ruta.docx"             # DOCX semanal (modo principal)
python main.py --docx "ruta.docx" --limpiar-historial  # test sin afectar producción
python main.py                                  # PDFs diarios (fallback)
python main.py --sin-ojv                        # omite M2
python main.py --hasta N                        # detiene tras Módulo N
python main.py --silencio                       # solo resúmenes
```

### `ojv_remates.py` (v10.0+) — MOTOR OJV BASE
**NO reescribir desde cero.** Automatización OJV con Playwright + RapidFuzz.

### `causas_ojv.xlsx`
BD interna con dos hojas:
- **REFERENCIA**: 233 filas con mapeo tribunal → corte (actualizada 2026-03-14: +29° y 30° Santiago)
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

- **Run DOCX semanal (184 bloques):** ~$3 USD en tokens Sonnet 4.6
- **Tiempo M1:** ~5 min (184 llamadas API secuenciales)
- **Tiempo M2:** ~21 min (51 causas en OJV)
- **Tiempo total pipeline:** ~27 min

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
