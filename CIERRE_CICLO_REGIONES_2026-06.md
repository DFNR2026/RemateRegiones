# Cierre de Ciclo — Regiones (PyL) + Auditoría CBR y Saneamiento del Pre-filtro v2
**Fecha de cierre:** 10 de junio de 2026
**Estado Regiones:** Producción v2 saneada — auditoría CBR cerrada con veredicto definitivo, red de seguridad de descartes en producción, bug del pre-filtro corregido y validado.
**Relación con RM (Mercurio):** la mentoría arquitectónica RM→Regiones sobre el filtro CBR queda **CERRADA**. RM sigue en producción estable, sin cambios este ciclo.

---

## 1. RESUMEN EJECUTIVO

Este ciclo cerró la mentoría CBR con un veredicto definitivo basado en datos reales con
cobertura completa, y en el camino la auditoría destapó y corrigió problemas más grandes
que el que se buscaba: una ruta v1 obsoleta que confundía el mapa de producción, un punto
ciego del 29% en el propio harness de auditoría, y un bug en el pre-filtro de producción v2
que descartaba causas válidas en silencio.

**Hechos clave:**
- Producción real de Regiones es **v2** (`v2_experimental/modulo1_v2.py`, Regex+Haiku,
  ~USD 0.05/corrida). La ruta v1 (`parsear_docx_semanal`, Sonnet, ~USD 3.00/corrida) está
  **obsoleta** y pendiente de eliminación (Tanda D).
- Veredicto CBR definitivo: **la rama de desambiguación ambigua NO muerde** (0 exclusiones
  sobre año ambiguo en 106/106 bloques; 0 en la población que producción evalúa).
  El fix asimétrico queda **archivado**. Vigilancia mensual de una línea.
- **3 commits** en `RemateRegiones`: `599131a` (harness), `e0bba23` (red de seguridad B1),
  y el commit B2 (pre-filtro superconjunto + sync réplica harness).
- **Decisión de negocio confirmada por el abogado/Diego:** las causas **A- (arbitrales) y
  V- (voluntarias) NO entran al alcance** de Regiones. Siguen otro camino procesal que no
  interesa. Su descarte silencioso en capa 2 es comportamiento correcto por diseño.

---

## 2. ESTADO DEL PIPELINE REGIONES (a hoy)

Ruta de producción: `python main.py --docx "ruta.docx"` (default = v2).

**Flujo v2 (`parsear_docx_v2`):**
1. PASO 2: segmentación por párrafos del DOCX (fechas de encabezado, títulos de semana).
2. PASO 3: expansión de párrafos con múltiples causas vía `separar_bloques()`
   (con rama `else` que conserva el párrafo original si no hay separación).
3. **PASO 3.5: pre-filtro `_RE_TIENE_ROL`** — descarta gratis (sin API) bloques sin patrón
   de ROL. Desde este ciclo: **los descartes se vuelcan a
   `logs/descartes_prefiltro_v2_*.csv`** (fecha_publicacion + primeros 300 chars) con
   `log.warning` — cola de revisión manual, ya no pérdida silenciosa.
4. `parsear_bloque_v2`: extraer_rol (regex, solo C-) → historial (dedup pre-API) → CBR
   (`evaluar_antiguedad_cbr` de `filtro_cbr`, compartido con v1) → tribunal regex con
   fallback Haiku → filtros RM / partidor.
5. Filtros post-parse: año < 2018, Banco Estado.

**Dependencias compartidas:** v2 importa 11 funciones de `modulo1_parser.py`
(`extraer_rol`, `separar_bloques`, `buscar_corte`, etc.) — ese archivo sobrevive a la
limpieza de v1 como **biblioteca**, solo se eliminan sus puntos de entrada muertos.

---

## 3. LO QUE SE CERRÓ ESTE CICLO

### Auditoría CBR — veredicto definitivo (mentoría RM→Regiones CERRADA)
- Harness `test_cbr_docx.py`: instrumentación fiel verificada espejo a espejo contra
  `filtro_cbr.py` (normalización, ventana de ancla, selección de marca "vigente").
- Corrido sobre DOCX real `(8) RESUMEN REG. 18 AL 22 DE MAYO DE 2026`:
  **106 bloques, EXCLUIR 26 / MANTENER 50 / REVISAR 30**.
- **Métrica de oro = 0** (exclusiones sobre año ambiguo), tanto en el total como en la
  población `PASA_PREFILTRO_V2=True` (la que producción realmente evalúa).
- La rama peligrosa (desambiguación por "vigente") disparó **1 vez en 106** — caso Quilpué
  (años 2020 vs 1988): eligió 1988 (Registro de Propiedad) sobre 2020 (Registro de
  Documentos / plano) → decisión CORRECTA → MANTENER. Lección: "el año más cercano al
  ancla" habría fallado aquí; el discriminador real es QUÉ registro referencia el ancla.
- **Fix asimétrico: ARCHIVADO** (gateado por datos; los datos dijeron que no hace falta).

### Bug del harness corregido en el camino (punto ciego del 29%)
- El PASO 3 del harness carecía de la rama `else` del pipeline: cuando `separar_bloques()`
  retornaba lista vacía, el bloque se botaba en silencio. Resultado: el harness auditaba
  60 de 106 bloques (~24 inicios de aviso reales invisibles).
- Fix: espejo exacto del PASO 3 de producción. Re-auditoría con cobertura completa.
- Commit `599131a` (incluye además columna `PASA_PREFILTRO_V2` y métricas de riesgo).

### Tanda B1 — Red de seguridad en el pre-filtro v2 (pérdida silenciosa → cola visible)
- El PASO 3.5 descartaba bloques sin patrón de ROL con solo un contador. 14 bloques/semana,
  de los cuales ~8 eran avisos judiciales reales y **2-3 causas válidas de negocio**.
- Fix: volcado a `logs/descartes_prefiltro_v2_*.csv` + `log.warning`. Validado 14/14 contra
  el harness. Commit `e0bba23`.

### Tanda B2 — Pre-filtro superconjunto del extractor (bug de punto de miles)
- **Bug:** `_RE_TIENE_ROL` usaba `\d+` (sin puntos) mientras `extraer_rol` admite
  `[\d.\s]*`. ROLes con punto de miles (`C-2.243-2020`, `C-23.173-2019`) morían en el
  pre-filtro sin llegar jamás al extractor que sí los parseaba. v1 los capturaba; v2 los
  perdía. **Principio violado: un pre-filtro debe ser superconjunto del extractor que
  gatea.**
- Fix: `\d+` → `\d[\d.\s]*` en `modulo1_v2.py` + réplica sincronizada en el harness.
- **Validación por predicción de 3 números, los 3 exactos:**
  pre-filtro 14→12, BancoEstado 6→7 (entró C-2.243-2020, Banco del Estado),
  sin-ROL post-parse 85→86 (entró C-23.173-2019, filtrado RM).
  CSV de descartes: subconjunto exacto, salieron solo los dos con punto de miles.

---

## 4. DECISIONES DE NEGOCIO (confirmadas este ciclo)

- **Causas A- (arbitrales) y V- (voluntarias): FUERA DE ALCANCE, definitivo.** No sirven
  al negocio; siguen otro camino procesal. Implicación técnica: que el pre-filtro las deje
  pasar (`[CcVvAa]`) y `extraer_rol` (solo C-) las descarte en capa 2 es coherente — quedan
  reconocidas como ROL (no ensucian la cola de revisión) y descartadas por alcance.
  NO se modifica el regex para extraerlas.
- **Pendiente de confirmar con el abogado:** si el CSV `descartes_prefiltro_v2_*.csv`
  (~12 bloques/semana) le sirve como cola de revisión manual en ese formato.

---

## 5. APRENDIZAJES CLAVE (transferibles)

- **Un pre-filtro debe ser superconjunto del extractor que protege.** Si el filtro barato
  es más estricto que el parser caro, se pierden casos que el parser habría resuelto — y
  se pierden en silencio, antes de cualquier log útil.
- **Un `else` no es cosmético.** `for sb in sub_bloques: append` vs
  `if len>1: ... else: append(original)` son programas distintos cuando la función puede
  retornar lista vacía. La "equivalencia en esencia" se demuestra con aritmética, no con
  prosa: 105−46+1=60 vs 105+1=106.
- **Validación por predicción:** fijar los números esperados ANTES de correr (14→12, 6→7,
  85→86) convierte una corrida en un experimento. Tres aciertos exactos ≈ entendimiento
  real del sistema; cualquier desvío = algo que no se entiende todavía.
- **Auditar el relato del agente contra el código, siempre.** Cinco veces en este ciclo la
  narrativa de Cline contradijo la evidencia ("se propagó automáticamente", "idéntica en
  esencia", anclas de diff cambiadas). Ninguna fue mala fe — todas fueron racionalizaciones
  plausibles que solo el verbatim desmintió. El costo de pedir verbatim es minutos; el
  costo de creerle al relato fue casi un KeyError en producción.
- **Los contadores con nombre engañoso son deuda activa:** `sin_rol` que cuenta
  historial+RM+CBR+sin-ROL no es un contador, es un cajón de sastre que esconde
  información en vez de revelarla.
- **Conocer la ruta de producción antes de instrumentar.** Una tanda completa (CSV en
  `parsear_docx_semanal`) se construyó y validó sobre una ruta obsoleta porque nadie
  preguntó primero "¿qué corre producción?". Costo recuperado con `git restore` (el
  congelamiento de commits hasta validar pagó), pero la pregunta va primero.
- **El punto ciego de un instrumento de medición invalida silenciosamente sus veredictos.**
  El harness con el bug del `else` habría certificado "la rama no muerde" sobre el 71% de
  la realidad. Auditar también al auditor.
- **Salvaguarda asimétrica (reafirmada):** ante la duda, REVISAR/visibilizar, no
  EXCLUIR/silenciar. Aplica al código (cola de revisión vs contador mudo) y al proceso
  (congelar commits hasta validar con datos reales).

---

## 6. PENDIENTES REGIONES

- **[TANDA C — parcialmente gateada]** Ampliación del regex de ROL para formatos perdidos
  que SÍ son de alcance:
  - Slash: `Rol C-5702/2024` (1 caso/semana observado; causa MANTENER por CBR = válida
    perdida). Fix de bajo riesgo: admitir `/` como separador antes del año en
    `_RE_TIENE_ROL` y `_RE_ROL`.
  - `Rol Nº 1840-2023` sin letra (5 casos/semana, incl. Santander y BCI con dominio
    antiguo). Riesgo de falsos positivos contra "rol de avalúo" — requiere guarda negativa
    y **medición offline contra el corpus de DOCX antes de tocar producción** (patrón
    harness). Nota a favor: los roles de avalúo observados (18-12, 140-42, 3900-84,
    1898-5) no terminan en 4 dígitos, el `-\d{4}` ya discrimina bastante.
  - Gateada por: confirmación del abogado sobre el formato de la cola de revisión.
- **[TANDA D]** Eliminar los puntos de entrada muertos de v1: `parsear_docx_semanal`,
  `parsear_diarios`, rama `--v1` en `main.py`. **NO borrar `modulo1_parser.py`** (es
  biblioteca de v2 y del harness: 11+ imports). Red: grep de imports antes del diff.
  Evaluar si v2 sale de `v2_experimental/` y pasa a ser el módulo titular.
- **[VIGILANCIA MENSUAL — CBR]** Con cada DOCX nuevo, correr
  `python test_cbr_docx.py "ruta.docx"` (gratis) y mirar UNA línea:
  `EXCLUSIONES SOBRE AÑO AMBIGUO (solo PASA_PREFILTRO_V2=True)`. Si se mantiene en 0
  durante 2-3 meses, cierre permanente. Si >0, reabrir el fix asimétrico archivado.
- **[ANOTADO, sin accionar]** Etiqueta `Sin ROL post-parse` en el resumen v2 es un cajón
  de sastre (mayormente historial + RM + CBR + partidor). Renombrar o desglosar por motivo
  cuando haya tanda de calidad de logs.
- **[ANOTADO, sin accionar]** `Haiku dir desactivado` deja direcciones en blanco mientras
  el contador `Sin dirección` marca 0. Confirmar si es configuración intencional.

---

## 7. PREFERENCIAS Y MÉTODO (para el nuevo chat)

- **Idioma:** español latinoamericano neutro ESTRICTO. NUNCA rioplatense (prohibido
  vos/tenés/podés/aplicá/mostrame/pasame). Usar tú/aplica/muéstrame/pásame.
- **Método:** PLAN (Cline solo lectura) antes de ACT. Un cambio / un foco / una validación
  con datos reales y **criterios numéricos fijados antes de correr**. Claude audita cada
  diff y log ANTES de aprobar; exige verbatim cuando el relato y los números no calzan.
  Commits congelados hasta validar. Nada de fixes a ciegas: si falta un archivo, se pide.
- **Herramientas:** Cline + DeepSeek V4 Pro implementan; Claude hace arquitectura, prompts
  PLAN/ACT y auditoría de diffs/logs. Diego corre los comandos.
- **Ecosistema:** RM (Mercurio, diario, scraping web, IA-primero) y Regiones (PyL, mensual,
  DOCX, regex-primero) comparten principios y contratos, NUNCA implementaciones copiadas
  sin adaptar. RM excluye lo que no es RM; Regiones excluye RM. A-/V- fuera de alcance
  en Regiones por decisión de negocio.
