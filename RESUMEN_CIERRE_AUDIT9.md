# Resumen de Cierre - Audit 9: Extracción y Cableado del Analizador de PDFs

## 1. Estado del Repositorio tras Commit f54cec2
- **Hito**: Frente B - Paso A1 (Extracción Verbatim del Analizador de PDFs).
- **Commit hash**: `f54cec2` (Sincronizado exitosamente con `origin/main`).
- **Módulos Nuevos**: `analisis_pdf.py` (206 líneas, UTF-8 nativo).
- **Métrica del Monolito**: Reducción de 4322 a **4129 líneas** (-193 líneas netas, acumulado -14.2% desde las 4813 líneas originales pre-Frente B).

## 2. Componentes Extraídos a `analisis_pdf.py`
El módulo fue aislado de forma verbatim respetando las barandas de seguridad:
- `_extraer_texto_pdf(ruta)` [L317-328 original] -> Lazy import de `log`.
- `_ocr_pdf(doc)` [L1386-1449 original] -> Lazy import de `log`. Preserva fixes de TESSDATA_PREFIX del Audit 6.1.
- `_analizar_pdf_acta(filepath)` [L1452-1559 original] -> Lazy import de `log`. Preserva `MONTO_MAXIMO_RAZONABLE` como variable de scope local.
- Constantes privadas del analizador: `_CARGO_REGEX` (L1380) y `_REGEX_MONTO_CLP` (L1383).

## 3. Resolución de Red Flags de Frontera
- **Helpers de Monto**: `_deuda_a_int` y `_parsear_monto` se determinaron como de uso global (14+ call sites). **Se quedan en el monolito**. El analizador no los requiere de forma directa en su lógica base actual.
- **Falso Positivo de Encoding**: Se verificó mediante scripts en frío que la "é" de `cr[eé]ditos?` está intacta en disco. El reporte previo fue un artefacto del pipe de la terminal cp1252 de Windows.
- **Filtro 2 Híbrido**: Se removió el separador estético contiguo de 3 líneas del monolito para garantizar la extracción limpia de las constantes.

## 4. Matriz de Validación y Smoke Test (100% PASS)
El cableado del filtrador se validó mediante el flujo híbrido en segundo plano pasando 5/5 pruebas críticas:
1. `py_compile filtrador_saldos.py` -> ✅ PASS. Sintaxis intacta.
2. `import filtrador_saldos` -> ✅ PASS. Runtime limpio.
3. Cross-import de dependencias -> ✅ PASS (log y _deuda_a_int resuelven nativos).
4. Re-export automático -> ✅ PASS. `from filtrador_saldos import _analizar_pdf_acta` responde transparente. **Protege el script WIP 'fix_rescate_actas.py' del Frente A contra contaminación.**
5. Conteo de líneas final -> ✅ PASS (4129 líneas).

El **Smoke Test operativo (`--solo-merge`)** procesó las 369 causas del Excel madre en producción conservando las 53 causas de 'Revisión Manual' y las 4 columnas extra del abogado (`_delta`, `_dias_desde_remate`, `observacion_abogado`, `fecha_revision_abogado`).

## 5. Pendientes para el Audit 10 (Frente A - Paso A2)
- El script de rescate `fix_rescate_actas.py` sigue congelado en su estado WIP.
- El caso testigo sigue siendo la causa **C-1855** con un excedente en riesgo de **$127,000,000**.
- El próximo paso requiere la inyección de las 3 nuevas reglas del analizador (Falso positivo por cláusula condicional de bases, extracción de montos posicionales alternativos, y restricción del criterio "es acta") directamente como un diff sobre `analisis_pdf.py`.