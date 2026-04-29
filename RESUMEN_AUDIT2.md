# RESUMEN AUDITORÍA 2 — 2026-04-14

## Resultados del abogado (22 causas revisadas)

### Decisión global: 16 ELIMINAR / 6 mantener

| Causa | Sección | Decisión | Observación abogado |
|-------|---------|----------|---------------------|
| C-1529-2023 | Excedente | **MANTENER** | OK, hay dos abogados, ubicar al demandado |
| C-2540-2025 | Excedente | **MANTENER** | Ya hablé con ella, es la matrona castigada |
| C-3276-2025 | Rev. Manual | **MANTENER** | Hablé con demandado, está en diálisis, llamar después de liquidación |
| C-3282-2025 | Rev. Manual | ELIMINAR | Avenimiento folio 24 Apremio |
| C-54-2024 | Rev. Manual | ELIMINAR | Juzgado de Pemuco no aparece en Corte de Chillán |
| C-592-2023 | Rev. Manual | ELIMINAR | Cargo al crédito folio 52 Apremio |
| C-992-2022 | Rev. Manual | ELIMINAR | Da cuenta de pago: capital + intereses + costas |
| C-1394-2019 | Rev. Manual | ELIMINAR | 3 abogados, posible avenimiento extrajudicial |
| C-53-2022 | Pend. Liq. | **MANTENER** | OK, ubicaré a los demandados (delta $123M) |
| C-1967-2023 | Pend. Liq. | ELIMINAR | Liquidación NEGATIVA, demandado quedó debiendo |
| C-1986-2019 | Pend. Liq. | ELIMINAR | Dos tercerías, no quedará saldo |
| C-5040-2024 | Pend. Liq. | ELIMINAR | Delta $1.1M muy bajo |
| C-953-2022 | Pend. Liq. | ELIMINAR | Deuda 2022, saldo muy bajo post intereses |
| C-2149-2024 | Pend. Liq. | ELIMINAR | Deudas adicionales Banco Estado |
| C-1485-2025 | Pend. Liq. | **MANTENER (SÍ)** | Probable excedente ~$6M, esperar liquidación |
| C-4837-2024 | Pend. Liq. | ELIMINAR | Pagaré con intereses altísimos |
| C-3601-2020 | Pend. Liq. | ELIMINAR | Avenimiento en cuaderno Principal folio 43 |
| C-621-2024 | Pend. Liq. | ELIMINAR | Suspensión de remate folio 21 |
| C-3628-2024 | Pend. Liq. | ELIMINAR | Renegociación extrajudicial |
| C-706-2025 | Pend. Liq. | ELIMINAR | Mandamiento ~$20M vs remate $25M |
| C-4306-2024 | Pend. Liq. | **MANTENER (SÍ)** | Acta $140M vs liquidación $96M, sí quedará saldo |
| C-850-2022 | Pend. Liq. | ELIMINAR | Remate suspendido, demandado con 3 abogados |

---

## Nuevos learnings (Auditoría 2)

### Confirmaciones de learnings previos
1. **Cargo al crédito** sigue apareciendo (C-592) — sistema lo detecta bien en F2 pero no en causas archivadas
2. **Da cuenta de pago** (C-992) — sistema lo detecta como señal informativa, abogado confirma = ELIMINAR
3. **Avenimiento** (C-3282, C-3601) — sistema lo detecta como señal informativa, abogado confirma = ELIMINAR
4. **Suspensión de remate** (C-621, C-850) — F1 detecta pero keyword liquidación lo trumpea

### Learnings nuevos
5. **Liquidación NEGATIVA** (C-1967) — existe el caso donde la liquidación muestra que el demandado quedó debiendo. F3 podría detectar esto sin riesgo.
6. **Tercerías** (C-1986) — terceros acreedores que absorben el saldo. No detectable automáticamente sin leer expediente completo.
7. **Deudas adicionales** (C-2149) — mismo deudor con múltiples juicios/créditos. Requiere búsqueda cruzada en OJV.
8. **Avenimiento en cuaderno equivocado** (C-3601) — avenimiento estaba en cuaderno Principal, no en Apremio. F1 solo lee Apremio.
9. **Intereses acumulados por antigüedad** (C-953, C-706, C-4837) — deudas antiguas (2019-2022) acumulan intereses que reducen/eliminan el delta aparente.
10. **Renegociación extrajudicial** (C-3628) — el demandado renegoció directamente con el banco fuera del expediente.
11. **Tribunal inubicable** (C-54) — Juzgado de Pemuco no aparece en la Corte de Chillán. Posible tribunal refundido.

### Análisis: ¿Qué podría automatizarse sin riesgo?

| Learning | ¿Automatizable? | Riesgo de falso positivo | Recomendación |
|----------|-----------------|-------------------------|---------------|
| Da cuenta de pago → ELIMINAR | Sí (promover señal) | MEDIO — pago parcial existe | NO por ahora |
| Avenimiento → ELIMINAR | Sí (promover señal) | MEDIO — avenimiento parcial | NO por ahora |
| Liquidación negativa | Sí (F3 regex) | BAJO | Candidato futuro |
| Tercerías | No sin leer expediente | N/A | Solo humano |
| Deudas adicionales | Parcial (búsqueda OJV) | ALTO | Solo humano |
| Delta mínimo threshold | Sí | MEDIO — depende de antigüedad | Evaluar umbral |
| Suspensión trumpea liquidación | Sí (cambiar prioridad F1) | MEDIO — remate puede reprogramarse | NO por ahora |

**Conclusión: No modificar filtros. La capa humana (22 causas → 6 mantener) funciona correctamente.**

---

## Estado post-auditoría 2 (después de ejecutar fix_eliminar_audit2.py)

| Métrica | Pre-audit2 | Post-audit2 | Delta |
|---------|-----------|-------------|-------|
| EXCEDENTE_CONFIRMADO | 2 | 2 | 0 |
| PENDIENTE_LIQUIDACION | 14 | 3 | -11 |
| REVISION_MANUAL | 6 | 1 | -5 |
| PENDIENTE_FILTRO1 | 169 | 169 | 0 |
| ELIMINAR | 50 | 66 | +16 |

### Causas activas post-audit2
| Causa | Estado | Delta | Nota abogado |
|-------|--------|-------|--------------|
| C-1529-2023 | EXCEDENTE | $61.9M | Ubicando demandado |
| C-2540-2025 | EXCEDENTE | $22.3M | Ya contactada |
| C-53-2022 | PEND_LIQ | $123.4M | Ubicando demandados |
| C-3276-2025 | REV_MANUAL | $10.7M | Demandado en diálisis |
| C-1485-2025 | PEND_LIQ | $13.7M | Esperar liquidación (~$6M excedente) |
| C-4306-2024 | PEND_LIQ | N/A* | Acta $140M vs liq $96M (~$44M excedente) |

*C-4306: el sistema no tenía monto_acta pero el abogado lo encontró manualmente.

### Excedente total perseguible: ~$84.3M confirmado + ~$50M potencial (C-4306 + C-1485)
