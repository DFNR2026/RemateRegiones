# PROYECTO: Sistema Automatizado de Análisis de Remates Judiciales - Chile
# Versión 7 del prompt maestro (actualizado 2026-03-20)

## CONTEXTO DEL NEGOCIO

Soy un inversionista inmobiliario chileno que analiza remates judiciales de propiedades para identificar oportunidades de alta rentabilidad. Proceso ~165 causas semanales del Diario P&L. Necesito automatizar todo el flujo.

**Mi flujo actual:**
1. Recibo el DOCX semanal consolidado del diario legal (Diario P&L)
2. El sistema parsea automáticamente cada aviso de remate
3. Busca la causa en la Oficina Judicial Virtual (OJV) del Poder Judicial
4. Descarga el Mandamiento (tiene la deuda) o Bases de Remate
5. Extrae montos y genera un reporte Excel

**Dirección de negocio actual:** Foco en propiedades **fuera de la Región Metropolitana** (regiones).

---

## ARQUITECTURA: 4 MÓDULOS ACTIVOS + ORQUESTADOR

```
D:\Remates\
├── main.py                ← orquestador que encadena los módulos
├── modulo1_parser.py      ← parser v1 (Sonnet 4.6) — fallback con --v1
├── v2_experimental/
│   └── modulo1_v2.py      ← parser v2 (Regex+Haiku) — DEFAULT para --docx
├── modulo2_ojv.py         ← consulta OJV via Playwright + extracción litigantes
├── modulo3_extractor.py   ← extrae montos de deuda desde PDFs descargados
├── modulo5_reporte.py     ← reporte final Excel con formato condicional
├── ojv_remates.py         ← motor OJV base (v10.0) — usado por modulo2_ojv.py
├── causas_ojv.xlsx        ← BASE DE DATOS INTERNA (REFERENCIA 234 tribunales + historial CAUSAS)
├── config.py              ← claves API (Anthropic), rutas, constantes, CAUSAS_IGNORADAS
├── limpiar_cache.py       ← limpieza de caché antes de test runs
├── logs/                  ← un .log por ejecución (dual-logging TeeWriter)
├── Descargas/             ← mandamientos y bases descargados por módulo 2
├── Diarios/               ← PDFs del diario P&L (fallback, ya no es fuente principal)
└── Reportes/              ← reportes Excel generados
```

Los datos fluyen **en memoria** (listas/dicts Python) entre módulos.

### ⛔ MÓDULO 4 — PERMANENTEMENTE ABANDONADO
La tasación automatizada fue **descartada definitivamente**. Los comparables a nivel comuna son insuficientes para decisiones de inversión reales.

**NO implementar, NO reactivar, NO sugerir alternativas de tasación automatizada.**

**DOS archivos Excel con propósitos distintos:**
1. `causas_ojv.xlsx` → BD INTERNA permanente:
   - Hoja REFERENCIA: 233 tribunales oficiales → cortes de apelaciones (solo agregar si faltan)
   - Hoja CAUSAS: historial acumulativo de ROLes procesados
2. `Reporte_YYYY-MM-DD.xlsx` → OUTPUT NUEVO cada ejecución semanal

---

## MÓDULO 1: Parser DOCX/PDFs + Deduplicación

### Parser v2 (DEFAULT): Regex-Flow + Haiku

**Archivo:** `v2_experimental/modulo1_v2.py` — se activa automáticamente con `--docx`.

**Estrategia de costo:** ROL/tribunal/demandante se extraen con regex puro (0 API calls). Solo dirección/comuna usan Claude Haiku como fallback cuando regex falla (~50% de causas). Tribunal usa Haiku solo si ambos regex (v1+v2) fallan.

**Costo:** ~$0.05/run vs ~$3/run (v1) — ahorro ~60x.

**Pipeline v2:**
1. Pre-filtro: bloques sin ROL se descartan GRATIS (regex)
2. ROL: regex `C-XXXXX-YYYY`
3. Tribunal: `extraer_tribunal_texto()` (v1) → `_extraer_tribunal_v2()` (4 patrones extra) → Haiku fallback
4. Demandante: `_extraer_demandante_v2()` (bancos, caratulada, slash patterns)
5. Dirección/comuna: `extraer_direccion_comuna()` (v1 regex) → `_extraer_comuna_v2()` (artículos) → Haiku fallback
6. Validación dirección: longitud >10 + dígito o palabra numérica
7. Comuna post-limpieza: truncar "sector"/"localidad"/"lugar"
8. Tribunal → Corte: `buscar_corte()` (RapidFuzz, ordinal/city recovery)
9. Comuna fallback: extraer ciudad del `tribunal_norm` (REFERENCIA)
10. Filtros: Banco Estado, pre-2018, RM, partidores/árbitros

### Parser v1 (fallback con `--v1`): Claude Sonnet

**Archivo:** `modulo1_parser.py` — se activa con `--v1 --docx` o modo PDFs diarios.
Usa Claude Sonnet 4.6 para extracción de TODOS los campos.

### Fuente principal: DOCX semanal

**Input:** Un archivo `.docx` semanal consolidado del Diario P&L (ej: `RESUMEN_16_AL_20_MARZO.docx`)

**Procesamiento DOCX:**
1. Leer con `python-docx`, separar párrafos por encabezados de fecha
2. Cada párrafo no-vacío = un bloque de causa
3. Si un párrafo contiene 2+ ROLes, splitear en sub-bloques
4. Extraer `fecha_publicacion` del encabezado + año del título

### Fuente fallback: PDFs diarios (siempre usa v1 Sonnet)

**Input:** Archivos PDF en `D:\Remates\Diarios\` (uno por día)

**Output:** Lista de dicts:
```python
{
    "rol": "32342", "año": "2015",
    "corte": "C.A. de Santiago",
    "tribunal": "1º Juzgado Civil de Santiago",
    "direccion": "Pasaje San Vicente Nº 41, comuna de Maipú",
    "comuna": "Maipú",
    "demandante": "Banco Itaú Chile",
    "demandado": "Pérez",
    "region_rm": True,
    "fecha_publicacion": "16 MARZO 2026"
}
```

---

## MÓDULO 2: Consulta OJV + Extracción de Litigantes

**Base:** `ojv_remates.py` (v10.0) — motor existente que NO se debe reescribir desde cero.

**Input:** Lista de causas con (rol, año, corte, tribunal)

**Qué hace:**
- **Filtro blacklist**: causas en `CAUSAS_IGNORADAS` se saltan
- Abre OJV con Playwright (perfil en `D:\Remates\.playwright_profile`)
- Busca cada causa, selecciona tribunal con **RapidFuzz** (umbral 85)
- Detecta tipo de procedimiento:
  - **Aceptados:** Ejecutivo Obligación de Dar, Ley de Bancos, Desposeimiento
  - **Descartados:** liquidación, ordinario, partición, arbitral, arrendamiento, monitorio
- **Ejecutivo** → descarga MANDAMIENTO del cuaderno Apremio
- **Ley de Bancos** → descarga BASES DE REMATE del cuaderno Principal
- **Extracción litigantes DTE/DDO** desde pestaña `#litigantesCiv`

**Output:** PDFs en `Descargas/`, enriquece causas con tipo_procedimiento, descargado, ruta_pdf, demandante/demandado OJV

---

## MÓDULO 3: Extractor de Montos de Deuda

**Input:** PDFs descargados por Módulo 2

- **Mandamientos:** busca "pague la suma de", "capital adeudado", etc.
- **Bases de Remate:** busca "mínimo para las posturas", "precio mínimo", etc.
- UF → CLP usando API `mindicador.cl/api/uf` (caché + fallback $38.500)

**Output:** enriquece con `monto_deuda_clp` (int, CLP), `monto_original` (str)

---

## MÓDULO 5: Reporte Final

**`actualizar_historial(causas)`:** APPEND a hoja CAUSAS (deduplicación por ROL+AÑO, idempotente)

**`generar_reporte(causas)`:**
- Clasifica: EXCELENTE (<50%), BUENA (<70%), REGULAR (<85%), DESCARTAR (≥85%), SIN DATOS
- Ordena por ratio ascendente
- `Reporte_YYYY-MM-DD.xlsx` con 3 pestañas: Resumen, RM, Regiones

---

## ORQUESTADOR (main.py)

```
python main.py --docx "ruta.docx"                          # DOCX semanal (v2 Regex+Haiku por defecto)
python main.py --v1 --docx "ruta.docx"                     # DOCX forzando Sonnet (v1)
python main.py --docx "ruta.docx" --limpiar-historial      # test sin afectar producción
python main.py                                              # PDFs diarios (fallback, usa v1 Sonnet)
python main.py --sin-ojv                                    # omite M2
python main.py --hasta N                                    # detiene tras Módulo N
python main.py --silencio                                   # solo resúmenes
```

**Dual-logging:** `logs/ejecucion_YYYYMMDD_HHMMSS.log`

---

## HOJA REFERENCIA (234 tribunales)

| CORTE | TRIBUNAL |
|---|---|
| C.A. de Arica | 1º Juzgado de Letras de Arica |
| C.A. de Santiago | 1º Juzgado Civil de Santiago |
| C.A. de Santiago | 29º Juzgado Civil de Santiago |
| C.A. de Santiago | 30º Juzgado Civil de Santiago |
| C.A. de Santiago | Juzgado de Letras de Colina |
| ... (234 tribunales en total) |

**Regla RM:** "C.A. de Santiago" o "C.A. de San Miguel" = RM. Todo lo demás = Regiones.

---

## DEPENDENCIAS

```
pymupdf (fitz), pandas, openpyxl, playwright, numpy, requests, rapidfuzz, anthropic, python-docx
```

---

## INSTRUCCIONES PARA CLAUDE CODE

1. **NO reescribir ojv_remates.py desde cero.** Solo modificar funciones específicas.
2. **Flujo de datos en memoria**, no archivos intermedios. Solo dos Excels: causas_ojv.xlsx (BD) y Reporte (output).
3. **Montos SIEMPRE en pesos chilenos (CLP).** Convertir UF via mindicador.cl.
4. **M4 NO EXISTE.** No implementar tasación automatizada de ningún tipo.
5. **RapidFuzz** es el motor de fuzzy matching. Umbral 80 en M1, umbral 85 en OJV.
6. **Ordinal recovery**: si ORDINAL MISMATCH, re-buscar en subset filtrado por ordinal correcto.
7. **City recovery**: si CITY MISMATCH, re-buscar en subset filtrado por ciudad del PDF.
8. **`_extraer_ordinal()`** reconoce numéricos Y textuales (diccionario hasta "Trigésimo").
9. **Reporte separado por región:** Pestaña RM y Pestaña Regiones.
10. **config.py** contiene todas las claves, rutas, constantes, `CAUSAS_IGNORADAS`.
11. **`--limpiar-historial`** para tests. `limpiar_cache.py` antes de cada test run.
12. **Hoja REFERENCIA** (234 tribunales): solo agregar si faltan, nunca eliminar.
13. **DOCX semanal es la fuente principal.** PDFs diarios son fallback.
14. **v2 es el parser DEFAULT para DOCX.** Usa `--v1` para forzar Sonnet.
15. **Litigantes DTE/DDO** se extraen de OJV y sobrescriben datos parciales de M1.
16. **Encoding logs**: usar solo ASCII en f-strings de log (`->` no `→`). Consola Windows usa cp1252.
17. **Costos API v2:** ~$0.05/run (Haiku). **v1:** ~$3/run (Sonnet). v2 es 60x más barato.

---

## MÉTRICAS DE PRODUCCIÓN (run 2026-03-20, v2)

| Métrica | Valor |
|---|---|
| Bloques procesados | 184 |
| Causas nuevas | 63 |
| Documentos descargados | 58/63 (92%) |
| Montos extraídos | 58/58 (100%) |
| Tiempo M1 (v2) | ~30s |
| Tiempo total pipeline | ~22 min |
| Cortes DESCONOCIDA | 1/63 (1.6%) |
| Llamadas Haiku | ~30 (dirección + tribunal fallback) |
| Costo API M1 | ~$0.05 |
