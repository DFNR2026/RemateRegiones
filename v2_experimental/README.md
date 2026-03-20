# Módulo 1 v2 — Regex-Flow + Haiku

## Objetivo
Reducir costos API de ~$3/run (Sonnet para todo) a ~$0.05/run manteniendo la misma calidad.

## Estrategia: Regex-Flow + Haiku

### 1. Pre-filtro regex (costo: $0)
Bloques sin patrón `C-\d+-\d+` / `V-\d+-\d+` / `A-\d+-\d+` se descartan antes de cualquier procesamiento. En el último run, 115/184 bloques eran ruido — esto los elimina gratis.

### 2. Extracción regex de campos fáciles (costo: $0)
- **ROL**: `(?:Rol|ROL)\s*:?\s*(?:Nº\s+)?(?:N°\s+)?(C-\d+-\d+|...)` — ya estaba en v1
- **Tribunal**: 11 estrategias regex heredadas de v1 (`extraer_tribunal_texto`)
- **Demandante**: regex caratulados, "causa X con Y", banco/slash patterns
- **Mínimo**: regex `$X.XXX` y `X UF`

### 3. Claude Haiku SOLO para dirección (costo: ~$0.05/run)
La dirección del inmueble es el campo genuinamente difícil: está embebido en prosa libre, mezclado con la dirección del tribunal, fragmentado por formato multi-columna.

- Se usa `claude-haiku-4-5-20251001` (20x más barato que Sonnet)
- **Solo se llama cuando regex falla** — muchas direcciones se extraen con regex
- Prompt corto y focalizado (solo dirección + comuna)
- Validación anti-alucinación (palabras de la dirección deben aparecer en el texto)

### 4. Funciones compartidas (importadas de v1)
`buscar_corte()`, `cargar_referencia()`, filtros Banco Estado, deduplicación, ordinal/city recovery — todo importado directamente de `modulo1_parser.py` sin duplicar código.

## Archivos

| Archivo | Descripción |
|---------|-------------|
| `modulo1_v2.py` | Parser optimizado con regex-flow + Haiku |
| `test_comparar.py` | Comparador v1 vs v2: tabla de diferencias + métricas |

## Uso

```bash
# Solo v2 (rápido, sin comparación)
python v2_experimental/test_comparar.py --docx "SANDBOX_V2_EDGE_CASES.docx" --solo-v2

# Comparación completa v1 vs v2
python v2_experimental/test_comparar.py --docx "SANDBOX_V2_EDGE_CASES.docx"
```

## Métricas esperadas

| Métrica | v1 (Sonnet) | v2 (Regex+Haiku) |
|---------|-------------|------------------|
| Llamadas API | N (una por bloque con ROL) | ~N/3 (solo regex fails) |
| Modelo | claude-sonnet-4-6 | claude-haiku-4-5-20251001 |
| Costo/call | ~$0.04 | ~$0.001 |
| Costo/run (~70 causas) | ~$3.00 | ~$0.05 |
| Campos: ROL, tribunal | Idénticos | Idénticos (mismo regex) |
| Campo: dirección | Baseline | ~90%+ match esperado |

## Trade-offs

- **v2 no extrae**: `demandado`, `fecha_remate` (no críticos para el pipeline downstream)
- **Dirección**: regex cubre ~70% de casos; Haiku cubre el resto
- **Tribunal**: regex de v1 ya es muy robusto (11 estrategias), no necesita LLM
