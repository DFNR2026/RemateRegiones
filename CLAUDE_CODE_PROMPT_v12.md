# PROYECTO: Sistema Automatizado de Análisis de Remates Judiciales - Chile
# Versión 12 del prompt maestro (actualizado 2026-04-11)

## CONTEXTO DEL NEGOCIO

Inversionista inmobiliario chileno que analiza remates judiciales para identificar excedentes (cuando la subasta supera la deuda). Proceso ~200 causas semanales del Diario P&L. Foco: propiedades fuera de la Región Metropolitana.

---

## ARQUITECTURA: DOS SISTEMAS INDEPENDIENTES

### Sistema 1: Pipeline Principal (M1-M5) ✅ EN PRODUCCIÓN

```
D:\Remates\
├── main.py                ← orquestador M1→M2→M3→M5
├── v2_experimental/modulo1_v2.py  ← parser v2 EN PRODUCCIÓN (Regex + Haiku)
├── modulo2_ojv.py         ← consulta OJV via Playwright, 5 workers subprocess
├── modulo3_extractor.py   ← extrae montos de deuda desde PDFs
├── modulo5_reporte.py     ← reporte final Excel
├── ojv_remates.py         ← motor OJV base (v10.0) — NO REESCRIBIR
├── causas_ojv.xlsx        ← BD INTERNA (235 tribunales + historial)
├── config.py              ← claves API, rutas, constantes
├── Regiones_docxToExcel.bat  ← doble-click: M1→M2→M3→M5
└── Reportes/              ← salida M5 = entrada del Filtrador
```

### Sistema 2: Filtrador de Saldos ✅ F1+F2+F3 FUNCIONALES

```
D:\Remates\
├── filtrador_saldos.py    ← ~4500 líneas, 5 workers subprocess
├── Detector_Excedentes.bat   ← doble-click: merge+F1+F2+F3
├── test_m2_sandbox.py     ← test M2 sin M1 (sin costo API)
├── temp_workers/          ← JSONs temporales workers filtrador
├── temp_workers_m2/       ← JSONs temporales workers M2
├── audit_html/            ← outerHTML tabla Apremio
├── Descargas/actas/       ← PDFs de actas descargados por F2
├── logs liqui/            ← logs del filtrador
└── Causas con liq/
    ├── Causas_posible_saldo.xlsx       ← Excel madre (BD viva)
    ├── Causas_posible_saldo_DD_mes_AAAA.xlsx  ← snapshots por fecha
    ├── Causas_con_liquidacion.xlsx     ← excedentes confirmados + rev manual
    ├── liquidaciones/                  ← PDFs de "Liquidacion (Credito)"
    └── liquidaciones_raw/             ← texto extraído de cada PDF
```

### ⛔ MÓDULO 4 — PERMANENTEMENTE ABANDONADO

---

## BYPASS WAF DEL PODER JUDICIAL (CRÍTICO)

PDJ tiene un WAF F5 BIG-IP que detecta y bloquea bots. La combinación que funciona es:

```python
# En TODOS los módulos que usan Playwright:
from playwright_stealth import Stealth

context = p.chromium.launch_persistent_context(
    profile_dir,
    headless=False,
    slow_mo=100,
    channel="chrome",                                    # Chrome real, no Chromium
    args=["--disable-blink-features=AutomationControlled"],
    accept_downloads=True,
)
page = context.pages[0] if context.pages else context.new_page()
Stealth().apply_stealth_sync(page)                       # Parchea fingerprint
```

**Reglas WAF:**
- `launch_persistent_context` reemplazó a `chromium.launch` — NO revertir
- `channel="chrome"` es obligatorio — Chromium bundled es bloqueado
- `playwright-stealth` es obligatorio — sin él, `consultaunificadacausas.php` rechaza
- Perfiles `.chrome-profile-wN` deben limpiarse (`shutil.rmtree`) al inicio de cada run para no cachear rechazos
- `navegar_a_consulta()` va a `home/index.php` → click "Consulta causas" → select competencia "Civil" → espera opciones de Corte
- URL anterior `indexN.php` ya NO se usa (redirige a login que el WAF bloquea)
- PDJ es históricamente inestable — si falla todo, verificar con abogados antes de asumir bug propio
- Ver `INCIDENTE_WAF_2026-04-09.md` para historial completo

---

## FLUJO SEMANAL

```
1. Llega DOCX semanal → D:\Remates\Diarios\
2. Doble-click Regiones_docxToExcel.bat → M1(~1min) → M2(5 workers, ~3min) → M3 → M5
3. Doble-click Detector_Excedentes.bat → F1+F2+F3(5 workers) → actualiza Excel madre
4. Si hay causas fallidas: python filtrador_saldos.py --reaudit --workers 5
5. Revisar resultados → generar DOCX auditoría para abogados
```

---

## WORKERS M2 (modulo2_ojv.py)

- 5 workers subprocess, distribución **round-robin** (no por Corte)
- Cada worker: `launch_persistent_context` con perfil `.chrome-profile-wN`
- Limpieza automática de perfiles al inicio de `_ejecutar_paralelo_m2()`
- Stagger de 3 segundos entre lanzamientos
- `procesar_causas_ojv(causas, n_workers=5)` — default desde main.py
- Fallback secuencial cuando `n_workers=1`
- `test_m2_sandbox.py N W` — testea N causas con W workers sin correr M1

---

## FILTRADOR DE SALDOS — DOCUMENTACIÓN COMPLETA

### Objetivo
Buscar excedentes post-remate. Documento clave: la **Liquidación** (~14 días post-remate).

### Vocabulario judicial
- Ejecutado = Demandado | Ejecutante = Demandante (banco)
- Acta de Remate = monto de adjudicación | Liquidación = saldo final
- Cargo al crédito = banco se adjudicó sin plata real → sin excedente → ELIMINAR

### Ejecución
```
python filtrador_saldos.py --workers 5                    # run normal
python filtrador_saldos.py --reaudit --workers 5          # reprocesar pendientes
python filtrador_saldos.py --recheck-cargo                # re-analizar PDFs locales
python filtrador_saldos.py --skip-pdf --workers 5         # sin descarga PDFs
python filtrador_saldos.py --solo-filtro3 --workers 5     # solo F3
python filtrador_saldos.py --recheck-liq                  # re-analizar liquidaciones locales
Detector_Excedentes.bat                                   # doble-click: merge+F1+F2+F3
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
NO buscar "Ordena liquidar" ni "Pone en conocimiento".

**Dos tipos de liquidación:**

| Tipo | Contenido PDF | Regex | Decisión |
|------|--------------|-------|----------|
| A | Incluye precio de remate + cálculo de saldo | `Saldo a favor del (Ejecutado\|Demandado)` | saldo > 0 → EXCEDENTE_CONFIRMADO |
| B | Solo calcula deuda total (sin comparar remate) | `(Crédito\|Capital) adeudado al` | delta = acta - crédito → según resultado |

### Estados del Excel madre
- `PENDIENTE_FILTRO1` — sin procesar
- `PENDIENTE_ACTA` — sin acta ni liquidación
- `NECESITA_PDF_ACTA` — acta detectada, PDF pendiente
- `PENDIENTE_LIQUIDACION` — posible excedente, esperando F3
- `PENDIENTE_REVISION_MANUAL` — requiere intervención humana
- `EXCEDENTE_CONFIRMADO` — saldo > 0, verificado
- `ELIMINAR` — sin excedente

### Señales informativas (no cambian estado, solo detalle)
- "giro cheque" / "gírese" / "consigna precio de remate"
- "objeta liquidación" / "nulidad de lo obrado"
- "da cuenta de pago" / "abandono procedimiento" / "desistimiento"

---

## RESULTADOS ACUMULADOS (2026-04-10, post reaudit)

| Métrica | Valor |
|---|---|
| Causas en Excel madre | ~270 |
| EXCEDENTE_CONFIRMADO | 2 (~$84.3M total) |
| PENDIENTE_LIQUIDACION | 14 |
| PENDIENTE_REVISION_MANUAL | 6 |
| PENDIENTE_ACTA | 6 |
| PENDIENTE_FILTRO1 | 169 |
| ELIMINAR | 50 |

### Excedentes confirmados
| Causa | Tipo | Excedente | Detalle |
|---|---|---|---|
| C-1529-2023 | A | $61,986,279 | Saldo a favor del ejecutado |
| C-2540-2025 | B | $22,385,979 | acta $61.2M - crédito $38.8M |

---

## DOCX AUDITORÍA PARA ABOGADOS

Archivo generado manualmente (no automatizado en bat). Diego sube Excel, Claude genera DOCX.

**Tres secciones:**
1. **Verde (#2E7D32):** Excedentes Confirmados
2. **Azul (#1F4E79):** Revisión Manual
3. **Naranja (#E65100):** Pendientes de Liquidación

Cada ficha: tabla unificada (label gris / valor blanco) + campos vacíos del abogado (seguimiento, observaciones, archivos).

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

### Reglas WAF / Playwright
25. **launch_persistent_context + channel="chrome" + playwright-stealth** — obligatorio en TODO módulo
26. **Limpiar perfiles .chrome-profile-wN** al inicio de cada run
27. **navegar_a_consulta()** va a `home/index.php` → click "Consulta causas" → select "Civil"
28. **NO usar `indexN.php` ni `consultaunificadacausas.php`** como URL directa

### Reglas de Claude (NO CC)
29. **SIEMPRE pedir los archivos .py antes de analizar o modificar.** Claude NO tiene acceso al sistema de archivos de Diego. Si necesita ver código, debe pedir que Diego suba el archivo. NUNCA suponer, imaginar o inventar el contenido de un archivo.
30. **Si Diego describe un error, pedir el log completo o screenshot** antes de diagnosticar. No adivinar.
31. **Prompts para CC como archivos .md separados** cuando el cambio es complejo (>20 líneas de instrucciones).

---

## GOTCHAS CONOCIDOS

- **`--limpiar-historial` activa parser v1 (PDFs)**, no v2 (DOCX). Para re-testear M2 con el mismo DOCX, limpiar manualmente las filas en `causas_ojv.xlsx` hoja CAUSAS.
- **PDJ inestable**: si todo falla, verificar con abogados si el sitio está caído a nivel nacional antes de buscar bugs en el código.
- **Reaudit recupera causas fallidas**: `--reaudit --workers 5` reprocesa PENDIENTE_ACTA sin perder trabajo previo.
- **`test_m2_sandbox.py`**: lee del reporte Excel, mapea cortes automáticamente. Usar para testear M2 sin costo API.
- **Desbalance workers filtrador**: distribución por Corte puede dejar W2 (Valparaíso) con el doble de causas. Round-robin no aplica porque se pierde optimización de tribunal agrupado.

---

## DEPENDENCIAS

```
pymupdf (fitz), pandas, openpyxl, playwright, numpy, requests, rapidfuzz,
anthropic, python-docx, pytesseract, Pillow, playwright-stealth
```

Requiere Tesseract-OCR instalado en Windows con idioma español (spa).
Requiere Google Chrome instalado (usado via `channel="chrome"`).
