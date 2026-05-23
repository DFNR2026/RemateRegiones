# Extrapolación: eliminar el fallback IA para dirección en M1 (proyecto Región Metropolitana / Sonnet)

**Origen:** validado en el proyecto de Regiones (`RemateRegiones`, parser v2 con Haiku) el 2026-04-29.
**Destino:** proyecto hermano de Región Metropolitana, que usa Sonnet (más caro por llamada → mayor ahorro).
**Estado:** CAMBIO PRINCIPAL probado y recomendado. FASE 2 solo documentada (no aplicar sin medir).

---

## Contexto y diagnóstico

En el pipeline M1 (parser del DOCX/PDF semanal), la **dirección del inmueble** se extrae primero con regex y, si el regex falla, se delega a un modelo de IA (Haiku en Regiones, Sonnet en RM) como fallback.

**Hallazgo clave:** la dirección es un campo **informativo**. Solo aparece en el DOCX que se le envía al abogado. **No la consumen** los módulos posteriores:
- M2 (OJV) busca por ROL + corte + tribunal.
- M3 (extracción de montos) usa el PDF de la causa.
- M5 (reporte) y el filtrador de saldos no dependen de la dirección.

Por lo tanto, pagar una llamada de IA para rellenar un campo que nadie usa aguas abajo no se justifica. El regex extrae la dirección cuando puede (direcciones urbanas con número); cuando no puede (rurales sin número: "sector Lomas Coloradas", "Cañal Bajo", "Alto Hospicio"), el campo queda vacío. Eso es aceptable.

### Impacto medido en Regiones (Haiku)

- Antes: la IA se llamaba en ~50% de las causas (~32% por dirección, ~19% por tribunal).
- Después de eliminar el fallback de dirección: **0 llamadas por dirección**, **-58% llamadas totales** (medido sobre 481 causas reales, 4 DOCX).
- **Comunas intactas** (ver sección "Preservar la comuna").

En el proyecto RM el ahorro debería ser mayor en términos de costo, porque Sonnet es más caro por llamada que Haiku.

---

## CAMBIO PRINCIPAL (probado, alto impacto) — eliminar el fallback IA de dirección

### Qué se hizo en Regiones (referencia)

En la función de parseo de bloque de `v2_experimental/modulo1_v2.py`, el flujo de dirección/comuna era:

1. `direccion, comuna = extraer_direccion_comuna(bloque)` — regex inicial.
2. `_extraer_comuna_v2(bloque)` — regex de comuna independiente (prefiere la versión más larga).
3. Validación: si la dirección es "corta/sin número", se descarta (`direccion = ""`).
4. **Bloque eliminado:** `if not direccion:` → llamaba a la IA (`_extraer_direccion_haiku`), validaba anti-alucinación, y además recuperaba la comuna (`if not comuna and com_h: comuna = com_h`).
5. Fix 12: si la comuna sigue vacía, se extrae la ciudad del nombre del tribunal normalizado.

El cambio reemplazó el paso 4 por: si el regex no da dirección válida, se deja vacía y se loguea. La función de IA (`_extraer_direccion_haiku`) y su prompt quedaron en el archivo **como referencia, sin invocar**, con un comentario `# DESACTIVADO [fecha]:`.

### Cómo identificar el bloque equivalente en el proyecto RM

Buscar en el código de M1 del proyecto RM (probablemente un `modulo1*.py` o equivalente):

1. La línea donde se llama al modelo Sonnet para dirección. Pistas de búsqueda:
   - Nombres de función tipo `_extraer_direccion_*`, `_direccion_ia`, `_direccion_sonnet`, `_direccion_claude`.
   - El `model=` con un identificador de Sonnet (ej. `claude-sonnet-...`).
   - Un prompt que pida "direccion" y "comuna" en JSON.
2. El bloque condicional que la invoca, típicamente `if not direccion:` después de la validación del regex.
3. Verificar si ese bloque **también recupera la comuna** (algo como `if not comuna and com_ia: comuna = com_ia`). **Esto es crítico** — ver sección siguiente.

### Cómo desactivarlo sin romper el esquema del dict de retorno

- **Eliminar / comentar** únicamente la invocación del modelo y su post-procesamiento de dirección.
- **Mantener** las claves `"direccion"` y `"comuna"` en el dict de retorno. La dirección simplemente vendrá vacía con más frecuencia. NO eliminar los campos: M2-M5 y el generador del DOCX esperan el esquema completo.
- Dejar la función de IA y su prompt en el archivo, sin invocar, con un comentario claro `# DESACTIVADO [fecha]: direccion informativa, no justifica costo IA`. Así la decisión es reversible y queda como referencia.
- Reformular el log: en vez de "regex no encontró dirección, llamando IA", poner algo como "Direccion regex vacia -- se deja en blanco (IA dir desactivada)". Mantener encoding ASCII en logs (Windows cp1252).
- Si el bloque de IA actualizaba algún contador de llamadas (`_api_calls` o similar), ese contador bajará solo al no invocar el modelo. Si hay un flag por causa tipo `_ia_used` que no se consume en ningún lado, puede quedar vestigial (verificar que no rompa nada).

### Preservar la comuna (CRÍTICO — verificar antes de aplicar)

El bloque de IA, además de la dirección, a veces recuperaba la **comuna**. Al eliminarlo, la comuna queda dependiendo de sus fuentes sin IA. En Regiones esas fuentes son:

1. `extraer_direccion_comuna` (regex).
2. `_extraer_comuna_v2` (regex).
3. **Extracción del nombre del tribunal** (el "Fix 12": si la comuna sigue vacía, se toma la ciudad del tribunal normalizado, ej. "…de La Serena" → "La Serena").

La fuente #3 es la red de seguridad clave: corre **después** del bloque de IA, así que cuando se elimina la IA, los casos sin comuna caen ahí. Como el tribunal casi siempre está presente, la comuna se recupera en la gran mayoría.

**Acción para RM antes de aplicar:**
- Confirmar que el proyecto RM tiene un equivalente a esa extracción de comuna desde el tribunal. **Si NO lo tiene**, la comuna podría degradarse más al quitar la IA → en ese caso, primero portar esa lógica (extraer ciudad del tribunal normalizado) y recién después desactivar la IA.
- Validar empíricamente: correr M1 standalone sobre un DOCX/lote ya procesado y comparar la tasa de comunas vacías contra un log anterior. Si las comunas no se degradan significativamente, el cambio está OK.

### Validación sugerida en RM

1. Correr M1 standalone (sin gatillar OJV) sobre un lote ya procesado.
2. Verificar en el log:
   - Desaparecen las líneas de "llamando IA / Sonnet" por dirección (solo deberían quedar las de tribunal, si existe ese fallback).
   - El conteo de llamadas API baja de forma notable.
   - Las comunas se siguen llenando (regex + tribunal), no quedan masivamente vacías.
   - Las direcciones con número siguen extrayéndose por regex.
3. Comparar contra un log anterior para cuantificar la reducción.

---

## FASE 2 OPCIONAL (documentar, NO aplicar sin medir) — huecos del regex de tribunal

> **No tocar el regex de tribunal todavía.** Esto queda anotado para evaluar solo si el volumen del proyecto Sonnet justifica el esfuerzo de medición.

El regex de tribunal tiene huecos en ciertas fórmulas de encabezado de los avisos. Detectados con código real en Regiones:

1. **"REMATE. [Tribunal] ordenó…"** — la estrategia de corte del nombre del tribunal corta con otras palabras clave pero **no con "ordenó"**, así que en esa fórmula el nombre del tribunal se captura mal o de más.
2. **"seguido ante" (masculino)** — el regex solo busca la variante femenina **"seguida"**, así que pierde los encabezados que usan "seguido ante [Tribunal]".

Cuando el regex de tribunal falla, hoy se delega al fallback de IA (que SÍ se mantiene, porque identificar mal el tribunal rompe la búsqueda en OJV). Es decir: estos huecos no producen errores, pero **sí generan llamadas de IA evitables**.

**Por qué no se optimiza ahora:** para mejorar el regex de tribunal con confianza se necesitan los **bloques de texto crudo** de las causas que efectivamente caen al fallback de IA (para ver las variantes reales de encabezado y no romper los casos que ya funcionan). Sin esa muestra, tocar el regex es riesgoso.

**Cuándo aplicar en RM:** solo si el volumen del proyecto Sonnet hace que ese ~19% de llamadas por tribunal sea costoso. En ese caso:
1. Instrumentar M1 para volcar a un archivo los bloques crudos de las causas que van al fallback de tribunal.
2. Acumular una muestra (varios lotes).
3. Ampliar el regex para cubrir "ordenó" como palabra de corte y "seguido"/"seguida" (ambos géneros).
4. Re-medir el % de llamadas de IA por tribunal antes/después.

---

## Checklist de aplicación en RM

- [ ] Localizar el bloque de fallback IA de dirección en el M1 de RM.
- [ ] Confirmar si ese bloque también recupera comuna.
- [ ] Confirmar que existe extracción de comuna desde el nombre del tribunal (si no, portarla primero).
- [ ] Desactivar la invocación de IA para dirección (dejar función + prompt como referencia, sin invocar).
- [ ] Mantener claves `"direccion"` y `"comuna"` en el dict de retorno.
- [ ] Reformular el log (ASCII).
- [ ] Validar M1 standalone: caída de llamadas IA + comunas no degradadas.
- [ ] (Opcional, Fase 2) Evaluar huecos del regex de tribunal solo si el volumen lo justifica.
