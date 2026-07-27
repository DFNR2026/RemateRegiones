# RESUMEN DE CIERRE — AUDIT7: Diagnostico rescate de actas + pacto flujo hibrido (Qwen local)

**Fecha de cierre:** 2026-05-26

Sesion larga que cubrio: (1) commit/push del fix OCR pendiente del Audit6.1, (2) diagnostico
completo del Frente A (actas escondidas bajo "Actuacion"), (3) construccion y dos dry-runs del
one-shot `fix_rescate_actas.py`, (4) deteccion de 3 bugs en el analizador de PDF que dejan el
Frente A parcial, y (5) el pacto de un NUEVO FLUJO DE TRABAJO HIBRIDO con Qwen3-Coder local.

---

## 0. CAMBIO DE FLUJO DE TRABAJO (lo mas importante de esta sesion)

A partir de ahora el desarrollo de Remates es HIBRIDO:

- **Claude (este chat) = Arquitecto de diseno y diagnostico.** Hace el analisis semantico pesado
  (logs cripticos, OCR raro, comportamiento de scraping, PDFs) y decide ESTRATEGICAMENTE el "que
  hacer". Entrega instrucciones ACOTADAS: firma exacta de funcion, que recibe/retorna, que NO
  tocar, y criterio de validacion. NO escribe scripts largos (quema cuota de Diego).
- **Qwen3-Coder:30b local (Roo Code/Ollama) = Ejecutor.** Pica la piedra: escribe modulos, maneja
  openpyxl, arma dry-runs. Gratis, ilimitado, privado, en la maquina de Diego (RTX 5070 Ti 16GB,
  Ryzen 5 9600X, 32GB DDR5). Es MoE (30B totales / 3.3B activos) -> corre fluido (~40-90 tok/s).
- **Diego = Arbitro.** Corre lo que Qwen produce, pega el output a Claude, Claude valida. El modelo
  local NUNCA decide si un resultado es correcto.
- **Claude Code (CC) queda en segundo plano**, se usa solo si Qwen se queda corto en algo puntual.
- **Regla de puntaje Qwen (1-10) OBLIGATORIA en cada tarea que Claude proponga:** 10 = mecanico/
  boilerplate/aislado, delegar a Qwen si o si. 1 = diagnostico criptico/analisis semantico, nefasto
  en local, volver a Claude/CC.

Existe un **`.clinerules` en `D:\Remates\`** (generado esta sesion) con las reglas del proyecto +
la estructura modular objetivo. Hace que las instrucciones de Claude a Qwen sean cortas (las
barandas ya estan puestas en el archivo).

---

## 1. COMMIT/PUSH DEL FIX OCR (cerrado)

Pendiente del Audit6.1, ahora resuelto. 5 commits en `origin/main`:
- `87612c8` fix OCR: TESSDATA_PREFIX + contador de paginas fallidas en _ocr_pdf
- `403021f` archivar fix_ocr_actas.py en scripts_oneshot_archivados/
- `410518c` eliminar EXTRAPOLACION...md (decision A de Diego)
- `14d2dc7` data: Excel madre + historial + BD interna + liquidaciones Excel
- `6ac0c29` ignorar actas_sin_monto.zip en .gitignore

`test_ocr_temp.py` borrado. Quedo PENDIENTE para Frente B (registrado por CC): politica de
binarios de liquidaciones (21 PDFs + 21 txt nuevos) y Reportes — decidir versionar todo vs
ignorar todo, y destrackear los 20 PDFs viejos para coherencia.

---

## 2. FRENTE A — diagnostico cerrado, ejecucion al ~70%

### Lo que esta RESUELTO y validado
- **Causa raiz:** las actas de remate existen en la tabla OJV pero bajo tramite generico
  "Actuacion" (no "Acta de remate"), por eso F2 no las detectaba (KEYWORDS_ACTA_REMATE solo busca
  el literal "acta de remate"). Confirmado con PDFs reales.
- **El rescate funciona:** el one-shot `fix_rescate_actas.py` encuentra el acta escondida
  escaneando candidatas genericas en una ventana de fechas y dejando que `_analizar_pdf_acta`
  decida cual es el acta. Dry-run probo: C-1855 (folio 52), C-2622 (folio 58), C-3259 (delta +$19.8M).
- **Gating ajustado (Cambio aplicado al one-shot):** ventana anclada al REMATE [remate-2, remate+14]
  (antes se anclaba a la senal de liquidacion, que dejaba fuera actas posteriores). Mas exclusion
  dura de filas con prefijo `[Nulo]`. Esto recupero C-2622 y bajo GATING de 9 a 5 (las 5 restantes
  son descartes correctos: nulo, dormidas, suspension pedida por el banco).

### Lo que FALTA (todo es el MISMO nudo: el analizador de PDF)
Los "4 bugs + C-1855 + bloqueo del --apply" son la misma cosa. Arreglar `_analizar_pdf_acta`
(idealmente ya extraido a `analisis_pdf.py`, Frente B) resuelve todo de golpe:

1. **Falso positivo de "cargo al credito" por clausula de bases.** El regex matchea "el ejecutante
   esta autorizado para adjudicarse con cargo al credito" (clausula condicional de las BASES, pag 1)
   en vez de la adjudicacion real. CASO TESTIGO: **C-1855 se marco ELIMINADA pero en realidad se
   adjudico a un TERCERO (Ana Maria Moreno Gomez) por $197M -> es un EXCEDENTE de ~$127M**
   (acta 197M - deuda 69.8M). Es el mas grave: hace perder excedentes reales.
2. **Monto mal extraido:** el regex agarra "la deuda" / "el minimo" / una caucion en vez de la
   adjudicacion. Ej C-2461 (agarro $77.2M de la deuda en vez de $68.8M de adjudicacion; decision
   final OK por casualidad porque igual es cargo al credito real).
3. **Criterio "es acta" demasiado laxo:** "monto>0" toma cualquier PDF con cifras como acta.
   Ej C-8848 f40 es un "Certificado de garantias suficientes" (lista de cauciones), no un acta;
   ademas f44 tiene "Suspension de remate" -> el remate no se hizo. Hay que (a) exigir ancla
   positiva de acta ("se adjudica"/"acta de remate"/"audiencia de remate") y (b) detectar
   suspension posterior al remate dentro de la ventana.

### Pasos para cerrar el Frente A (en orden)
1. Disenar las 3 reglas del analizador (TAREA DE CLAUDE, puntaje Qwen ~5: el diseno es de Claude,
   la escritura del regex/condicion la pica Qwen una vez definida).
2. Re-correr dry-run; validar que C-1855 pase a excedente ~$127M y los montos falsos se corrijan.
3. `--apply` con spot-check de PDFs de las ELIMINADA-cargo (irreversibles en el proximo run).
4. C-3512-2022 (deuda=$0): mini-fix aparte estilo fix_deuda_cero (puntaje Qwen 9, mecanico).
5. Integrar el rescate al filtrador para que corra en --reaudit, no a mano (puntaje Qwen 6).

### Estado del one-shot
`fix_rescate_actas.py` esta en `D:\Remates\` como **WIP, sin commitear, sin aplicar.** El gating
y el rescate funcionan; falta el analizador. NO correr --apply hasta arreglar el analizador
(si no, se elimina C-1855 y se pierden $127M).

### Validacion ya hecha (no repetir)
- Diagnostico contra HTML reales (audit_html) y PDFs reales abiertos por Diego.
- OCR replicado y confirmado: el fix OCR del Audit6.1 (TESSDATA_PREFIX) es lo que permite leer
  estas actas escaneadas. C-153 ($216M cargo), C-234 ($82M cargo) salen correctos.
- El spot-check pre-apply FUNCIONO: atrapo los 3 falsos positivos antes de escribir. Mantener
  ese paso siempre.

---

## 3. MAPA DE PENDIENTES REAL (estructurado con Diego)

### Frente A (Parcial, ~70%)
Los 3 bugs del analizador de PDF + C-3512-2022 (deuda=$0) + integrar el rescate al filtrador
+ desbloquear el --apply + resolver C-1855. (Todo es el mismo nudo del analizador.)

### Frente B (Modularizacion y limpieza) — ARRANCA MANANA
- Modularizar `filtrador_saldos.py` (~4800 lineas, no 900) segun la Regla 5 del .clinerules:
  persistencia_excel.py / analisis_pdf.py / tabla_historial.py / descarga_pdf.py /
  orquestador_workers.py / procesamiento.py / filtrador_saldos.py (entrypoint delgado).
  ojv_remates.py NO se toca.
- Renombrar `v2_experimental` (ya es produccion), marcar obsoletos (`modulo1_parser.py` solo se
  usa con --v1), actualizar a `CLAUDE_CODE_PROMPT_v13.md`.
- Politica de binarios de liquidaciones/Reportes (versionar todo vs ignorar todo + destrackear
  los 20 viejos).
- **Primer trabajo de Qwen (manana):** extraer `persistencia_excel.py` (puntaje 9, ideal para
  estrenar el flujo) + limpieza de obsoletos en `modulo1_parser.py`.

### Frente C (Bajos/bugs menores)
- `UnicodeEncodeError` en `_TeeWriter` (ya diagnosticado Audit5: sanitizar al escribir a
  sys.__stdout__ con errors='replace'). Puntaje Qwen 8.
- Los 5 PDFs que M3 no extrae (C-1528, C-2218, C-2831, C-3512, C-3676). Esto es DIAGNOSTICO
  (abrir PDFs, entender el patron que rompe) -> puntaje Qwen 3, traer a Claude.

### PAUSADO por completo
- **Extrapolacion al proyecto RM (Sonnet->Haiku).** Pertenece a otro repositorio. Se retoma al
  final, cuando aqui ya este todo aprendido. No mezclar en este directorio.
- **Claude Cowork:** eliminado de pendientes (decision de Diego).

---

## 4. PUNTAJES QWEN POR TAREA (referencia, evaluados esta sesion)

| Tarea | Puntaje Qwen | Quien ejecuta |
|---|---|---|
| Extraer persistencia_excel.py (Frente B) | 9 | Qwen |
| C-3512-2022 deuda=$0 mini-fix | 9 | Qwen |
| Renombrar v2_experimental / obsoletos | 8 | Qwen |
| UnicodeEncodeError _TeeWriter | 8 | Qwen |
| Politica de binarios (gitignore/destrackear) | 7 | Qwen (tras decision Claude+Diego) |
| Integrar rescate al filtrador | 6 | Qwen (Claude define el donde) |
| Arreglar 3 reglas del analizador de PDF | 5 | Mixto (diseno Claude, escritura Qwen) |
| Diagnostico 5 PDFs que M3 no extrae | 3 | Claude |

Patron: Frente B y fixes ya-diagnosticados de C son territorio Qwen (7-9). El analizador del
Frente A y los diagnosticos nuevos son de Claude (3-5).

---

## 5. CHECKLIST AL INICIAR PROXIMO CHAT

1. Subir este RESUMEN_CIERRE_AUDIT7.md.
2. Confirmar que Ollama + Qwen3-Coder:30b + Roo Code quedaron instalados (Diego era nuevo en esto).
3. Confirmar que el `.clinerules` esta en `D:\Remates\`.
4. Arrancar Frente B: Claude da la orden acotada para extraer `persistencia_excel.py`; Qwen ejecuta;
   Diego corre y pega el output; Claude valida.
5. Recordar: en cada tarea Claude da puntaje Qwen 1-10. El local nunca decide correctitud.
