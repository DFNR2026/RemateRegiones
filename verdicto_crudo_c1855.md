# Veredicto crudo de `_analizar_pdf_acta` — C-1855 f52 y f45

## Salida literal de `repr(_analizar_pdf_acta(ruta))`

```
rescate_1855_2025_f52.pdf -> {'cargo_al_credito': False, 'monto_adjudicacion': 197000000, 'texto_monto': '$197.000.000.-'}
rescate_1855_2025_f45.pdf -> {'cargo_al_credito': False, 'monto_adjudicacion': None, 'texto_monto': None}
```

## Línea exacta de `_CARGO_REGEX` en `analisis_pdf.py` (línea 11)

```python
_CARGO_REGEX = re.compile(r'cargo\s+(?:a\s+(?:su|sus|los?)|al)\s+cr[eé]ditos?', re.IGNORECASE)