# REPORTE DE INSPECCIÓN TÉCNICA - INGENIERÍA INVERSA PIPELINE
**Proyecto**: Filtro Automatizado de Antigüedad CBR (Regiones / RM)
**Fecha de Inspección**: 28 de Mayo de 2026
**Herramienta de Escaneo**: Cline (PLAN MODE)

---

## 1. Archivo y Función que Procesa el Texto Plano del Aviso

| Atributo | Detalle Técnico |
| :--- | :--- |
| **Archivo Encontrado** | `modulo1_parser.py` |
| **Función Crítica** | `parsear_bloque(bloque_raw: str, df_ref: pd.DataFrame, historial_roles: set | None = None) -> dict | None` |
| **Línea de Entrada Raw** | **Línea 920** (a través del parámetro `bloque_raw`) |
| **Línea de Sanitización** | **Línea 933** (`bloque = limpiar_texto(bloque_raw)`) |

---

## 2. Secuencia y Flujo de Control en `parsear_bloque`

El pipeline procesa los extractos siguiendo un orden estrictamente lineal antes de delegar en los LLM:

1. **Línea 936 - Filtro de Bloques Corruptos**: Detecta patrones de ruido crítico como `...!!!` en el string raw.
2. **Línea 938 - Filtro Preventivo de Árbitros**: Ejecuta expresiones regulares nativas para identificar liquidaciones de jueces partidores.
3. **Líneas 942-945 - Extracción de ROL**: Llama a `extraer_rol()`. Si el veredicto es `None`, aborta la secuencia inmediatamente.
4. **Líneas 948-951 - Validación de Historial de Procesamiento**: Verifica si el ROL ya fue procesado con anterioridad en la base de datos distribuida para saltarse llamadas redundantes.
5. **Línea 954 - Invocación Core LLM (Claude API)**: Ejecuta `extraer_campos_claude(bloque)`. Esta es la zona de alto impacto operativo (~$0.01 por ejecución).
6. **Líneas 978-982 - Filtro Post-LLM**: Valida si el tribunal determinado estructuralmente contiene palabras clave de particiones de herencia.
7. **Línea 1061 - Segmentación Regional**: Descarta o clasifica si la corte pertenece geográficamente a la Región Metropolitana (`corte in CORTES_RM`).

---

## 3. Punto de Inyección Quirúrgica (Filtro Antigüedad CBR)

Para maximizar la eficiencia y reducir a cero el consumo innecesario de tokens en propiedades nuevas que deban descartarse en la frontera, la ventana idónea de inyección es la **Línea 953**.

### Arquitectura de Bloque Propuesta:

```python
    # =========================================================================
    # ── FILTRO TRANSVERSAL: ANTES DE INVOCAR APIS (AHORRO OPERATIVO) ─────────
    # =========================================================================
    if not _evaluar_dominio_vigente(bloque, rol, anio):
        return None

    # ── Extracción principal: Claude API ─────────────────────────────────────
    campos = extraer_campos_claude(bloque)