# PROYECTO: Sistema Automatizado de Análisis de Remates Judiciales - Chile
# Versión 6 del prompt maestro (actualizado 2026-03-14)

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
├── modulo1_parser.py      ← parser DOCX/PDFs (Claude API Sonnet 4.6) + deduplicación + filtros
├── modulo2_ojv.py         ← consulta OJV via Playwright + extracción litigantes
├── modulo3_extractor.py   ← extrae montos de deuda desde PDFs descargados
├── modulo5_reporte.py     ← reporte final Excel con formato condicional
├── ojv_remates.py         ← motor OJV base (v10.0) — usado por modulo2_ojv.py
├── causas_ojv.xlsx        ← BASE DE DATOS INTERNA (REFERENCIA 233 tribunales + historial CAUSAS)
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

### Fuente principal: DOCX semanal (desde 2026-03-14)

**Input:** Un archivo `.docx` semanal consolidado del Diario P&L (ej: `RESUMEN_16_AL_20_MARZO.docx`)

**Estructura del DOCX:**
```
Párrafo: "RESUMEN NACIONAL"                         ← IGNORAR (título)
Párrafo: "REMATES SEMANA DEL 16 AL 20 MARZO 2026"   ← Extraer año
Párrafo: "16 MARZO"                                  ← Encabezado de fecha
Párrafo: "EXTRACTO Remate. 1º Juzgado..."           ← Una causa completa
Párrafo: (vacío)                                     ← IGNORAR
Párrafo: "17 MARZO"                                  ← Encabezado de fecha
...
```

**Motor de extracción:** Claude API (`claude-sonnet-4-6`). API key en `config.py`.

**Procesamiento DOCX:**
1. Leer con `python-docx`, separar párrafos por encabezados de fecha (regex: `r'^\d{1,2}\s+(ENERO|...|DICIEMBRE)$'`)
2. Cada párrafo no-vacío que no sea título ni fecha = un bloque de causa
3. Si un párrafo contiene 2+ ROLes, splitear en sub-bloques
4. Extraer `fecha_publicacion` del encabezado + año del título

### Fuente fallback: PDFs diarios

**Input:** Archivos PDF en `D:\Remates\Diarios\` (uno por día)
- Extrae texto con PyMuPDF, separa en bloques por marcadores

### Pipeline común (ambas fuentes)

5. Filtro jueces partidores/árbitros pre-extracción
6. Enviar bloque a Claude API: extraer ROL, tribunal, demandante, demandado, dirección, comuna
7. **`_limpiar_tribunal()`**: une guiones silábicos, elimina direcciones físicas, normaliza capitalización
8. Filtro partidor/árbitro post-extracción
9. Deduplicar entre bloques de la semana + contra historial CAUSAS
10. Filtros: Banco Estado, causas pre-2018
11. **Matching tribunal → corte con RapidFuzz** (`token_set_ratio`, umbral 80):
    - **Validación ordinal**: si ordinal no coincide → **recovery** (re-busca solo entre tribunales con ordinal correcto)
    - **`_extraer_ordinal()`**: reconoce numéricos (1°, 29°) Y textuales ("Séptimo"→7, "Vigésimo Noveno"→29)
    - **Validación de ciudad**: compara ciudades normalizadas (sin tildes). Si difieren → penaliza 0.7x → si cae bajo umbral → **city recovery** (re-busca solo entre tribunales con la ciudad del PDF)
12. **Retry dirección**: segundo intento con prompt focalizado si dirección=None
13. Clasificar región: C.A. de Santiago / C.A. de San Miguel = RM. Todo lo demás = Regiones.

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
python main.py --docx "ruta.docx"                          # DOCX semanal (principal)
python main.py --docx "ruta.docx" --limpiar-historial      # test sin afectar producción
python main.py                                              # PDFs diarios (fallback)
python main.py --sin-ojv                                    # omite M2
python main.py --hasta N                                    # detiene tras Módulo N
python main.py --silencio                                   # solo resúmenes
```

**Dual-logging:** `logs/ejecucion_YYYYMMDD_HHMMSS.log`

---

## HOJA REFERENCIA (233 tribunales)

| CORTE | TRIBUNAL |
|---|---|
| C.A. de Arica | 1º Juzgado de Letras de Arica |
| C.A. de Santiago | 1º Juzgado Civil de Santiago |
| C.A. de Santiago | 29º Juzgado Civil de Santiago |
| C.A. de Santiago | 30º Juzgado Civil de Santiago |
| ... (233 tribunales en total) |

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
12. **Hoja REFERENCIA** (233 tribunales): solo agregar si faltan, nunca eliminar.
13. **DOCX semanal es la fuente principal.** PDFs diarios son fallback.
14. **Litigantes DTE/DDO** se extraen de OJV y sobrescriben datos parciales de M1.
15. **Encoding logs**: usar solo ASCII en f-strings de log (`->` no `→`). Consola Windows usa cp1252.
16. **Costos API:** ~$3 por run semanal completo (184 bloques, Sonnet 4.6).

---

## MÉTRICAS DE PRODUCCIÓN (run 2026-03-14)

| Métrica | Valor |
|---|---|
| Bloques procesados | 184 |
| Causas nuevas | 55 |
| Documentos descargados | 51/55 (92.7%) |
| Montos extraídos | 51/51 (100%) |
| Tiempo total | ~27 min |
| Cortes DESCONOCIDA | 1/55 (1.8%) |
