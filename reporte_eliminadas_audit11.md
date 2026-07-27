# Reporte: ELIMINADAS re-analizadas con discriminador Audit11

**Fecha:** 2026-06-04T22:51:50.564455
**ELIMINADAS totales en Excel:** 11
**Subconjunto con 'cargo' / 'eliminada_cargo' en log_decision:** 6

## A. Listado completo de ELIMINADAS

| # | ROL | AÑO | Demandante | Demandado | Deuda | Monto Acta | Motivo (log_decision) |
|---|---|---|---|---|---|---|---|
| 1 | 8848 | 2024 | BANCO DE CRÉDITO E INVERSIONES | MARLON HANS VENEGAS FUENTES | 50190615 | 10000000 | RescateActa [2026-05-27]: Adjudicacion $10.000.000 < deuda $50.190.615 (folio 40). | REAUDIT: Sin senales nuevas |
| 2 | 153 | 2025 | BANCO DE CHILE | SARA ANDREA RAMÍREZ CRUZ | 83216929 | nan | RescateActa [2026-05-27]: Cargo al credito (folio 58). Sin excedente. | REAUDIT: Sin senales nuevas |
| 3 | 3459 | 2025 | BANCO SANTANDER-CHILE S.A | CRISTIAN PATRICIO SALAZAR MANC | 134509927 | nan | RescateActa [2026-05-27]: Cargo al credito (folio 21). Sin excedente. | REAUDIT: Sin senales nuevas |
| 4 | 4163 | 2024 | SCOTIABANK CHILES.A | MANUEL ALEJANDRO NÚÑEZ MORIS | 64646604 | nan | RescateActa [2026-05-27]: Cargo al credito (folio 65). Sin excedente. | REAUDIT: Senal 4: "Solicita liquidación" (folio 57, 2026-04-13) -> directo a Filtro 3 |
| 5 | 951 | 2024 | KÜPFER HERMANOS S.A. | JOCTAN ANDRÉS VILLARROEL IRRIB | 9390893 | 1500000 | RescateActa [2026-05-27]: Adjudicacion $1.500.000 < deuda $9.390.893 (folio 46). | REAUDIT: Senal 4: "Ordena liquidar el crédito" (folio 58, 2026-05-20) -> directo a Filtro 3 |
| 6 | 234 | 2025 | BANCO DE CHILE | FERNANDO ADOLFO ARCOS LEROUX | 1336958 | nan | RescateActa [2026-05-27]: Cargo al credito (folio 46). Sin excedente. | REAUDIT: Senal 4: "Ordena liquidar el crédito" (folio 48, 2026-05-18) -> directo a Filtro 3 |
| 7 | 2751 | 2021 | BANCO DE CHILE | INVERSIONES TRAUCA SPA | 15000000 | 3200000 | RescateActa [2026-05-27]: Adjudicacion $3.200.000 < deuda $15.000.000 (folio 76). | REAUDIT: Senal 4: "Ordena liquidar el crédito" (folio 84, 2026-05-15) -> directo a Filtro 3 |
| 8 | 1522 | 2025 | SCOTIABANK CHILE S.A | MARÍA JOSÉ MATURANA OTEY | 16202127 | 4000000 | RescateActa [2026-05-27]: Adjudicacion $4.000.000 < deuda $16.202.127 (folio 24). | REAUDIT: Sin senales nuevas |
| 9 | 5979 | 2025 | BANCO SANTANDER-CHILE S.A | JOSÉ ALEJANDRO HERNÁNDEZ QUILO | 45301317 | 3545885 | RescateActa [2026-05-27]: Adjudicacion $3.545.885 < deuda $45.301.317 (folio 19). | REAUDIT: Sin senales nuevas |
| 10 | 2461 | 2024 | BANCO SANTANDER-CHILE S.A. | ELISEO GUILLERMO SAAVEDRA PÉRE | 83341971 | nan | RescateActa [2026-05-27]: Cargo al credito (folio 31). Sin excedente. | REAUDIT: Sin senales nuevas |
| 11 | 2622 | 2024 | BANCO DE CRÉDITO E INVERSIONES | IGNACIO ANDRÉS GALLEGUILLOS PA | 97993864 | nan | RescateActa [2026-05-27]: Cargo al credito (folio 58). Sin excedente. | REAUDIT: Senal 4: "Ordena liquidar el crédito" (folio 56, 2026-05-18) -> directo a Filtro 3 |

## B. Re-análisis con discriminador NUEVO (causas con 'cargo')

### B1. ROL 153-2025 — BANCO DE CHILE

- **Demandado:** SARA ANDREA RAMÍREZ CRUZ
- **MONTO_DEUDA_CLP:** 83216929
- **monto_acta_remate:** nan
- **log_decision:** RescateActa [2026-05-27]: Cargo al credito (folio 58). Sin excedente.
REAUDIT: Sin senales nuevas
- **PDFs encontrados:** 1
  - `D:\Remates\Descargas\rescate_actas\rescate_153_2025_f58.pdf`
- **Veredicto NUEVO (crudo):** `adjudicatario='ejecutante'`, `cargo_al_credito=True`, `monto_adjudicacion=$216,000,000`

### B2. ROL 3459-2025 — BANCO SANTANDER-CHILE S.A

- **Demandado:** CRISTIAN PATRICIO SALAZAR MANCILLA
- **MONTO_DEUDA_CLP:** 134509927
- **monto_acta_remate:** nan
- **log_decision:** RescateActa [2026-05-27]: Cargo al credito (folio 21). Sin excedente.
REAUDIT: Sin senales nuevas
- **PDFs encontrados:** 1
  - `D:\Remates\Descargas\rescate_actas\rescate_3459_2025_f21.pdf`
- **Veredicto NUEVO (crudo):** `adjudicatario='ejecutante'`, `cargo_al_credito=True`, `monto_adjudicacion=$109,000,000`

### B3. ROL 4163-2024 — SCOTIABANK CHILES.A

- **Demandado:** MANUEL ALEJANDRO NÚÑEZ MORIS
- **MONTO_DEUDA_CLP:** 64646604
- **monto_acta_remate:** nan
- **log_decision:** RescateActa [2026-05-27]: Cargo al credito (folio 65). Sin excedente.
REAUDIT: Senal 4: "Solicita liquidación" (folio 57, 2026-04-13) -> directo a Filtro 3
- **PDFs encontrados:** 2
  - `D:\Remates\Descargas\C-4163-2024_MANDAMIENTO.pdf`
  - `D:\Remates\Descargas\rescate_actas\rescate_4163_2024_f65.pdf`
- **Veredicto NUEVO (crudo):** `adjudicatario='indeterminado'`, `cargo_al_credito=False`, `monto_adjudicacion=$87,423,498`

### B4. ROL 234-2025 — BANCO DE CHILE

- **Demandado:** FERNANDO ADOLFO ARCOS LEROUX
- **MONTO_DEUDA_CLP:** 1336958
- **monto_acta_remate:** nan
- **log_decision:** RescateActa [2026-05-27]: Cargo al credito (folio 46). Sin excedente.
REAUDIT: Senal 4: "Ordena liquidar el crédito" (folio 48, 2026-05-18) -> directo a Filtro 3
- **PDFs encontrados:** 4
  - `D:\Remates\Descargas\C-234-2025_MANDAMIENTO.pdf`
  - `D:\Remates\Descargas\rescate_actas\rescate_234_2025_f44.pdf`
  - `D:\Remates\Descargas\rescate_actas\rescate_234_2025_f45.pdf`
  - `D:\Remates\Descargas\rescate_actas\rescate_234_2025_f46.pdf`
- **Veredicto NUEVO (crudo):** `adjudicatario='indeterminado'`, `cargo_al_credito=False`, `monto_adjudicacion=$1,336,958`

### B5. ROL 2461-2024 — BANCO SANTANDER-CHILE S.A.

- **Demandado:** ELISEO GUILLERMO SAAVEDRA PÉREZ
- **MONTO_DEUDA_CLP:** 83341971
- **monto_acta_remate:** nan
- **log_decision:** RescateActa [2026-05-27]: Cargo al credito (folio 31). Sin excedente.
REAUDIT: Sin senales nuevas
- **PDFs encontrados:** 2
  - `D:\Remates\Descargas\C-2461-2024_MANDAMIENTO.pdf`
  - `D:\Remates\Descargas\rescate_actas\rescate_2461_2024_f31.pdf`
- **Veredicto NUEVO (crudo):** `adjudicatario='indeterminado'`, `cargo_al_credito=False`, `monto_adjudicacion=$77,208,723`

### B6. ROL 2622-2024 — BANCO DE CRÉDITO E INVERSIONES

- **Demandado:** IGNACIO ANDRÉS GALLEGUILLOS PALACIOS
- **MONTO_DEUDA_CLP:** 97993864
- **monto_acta_remate:** nan
- **log_decision:** RescateActa [2026-05-27]: Cargo al credito (folio 58). Sin excedente.
REAUDIT: Senal 4: "Ordena liquidar el crédito" (folio 56, 2026-05-18) -> directo a Filtro 3
- **PDFs encontrados:** 2
  - `D:\Remates\Descargas\C-2622-2024_MANDAMIENTO.pdf`
  - `D:\Remates\Descargas\rescate_actas\rescate_2622_2024_f58.pdf`
- **Veredicto NUEVO (crudo):** `adjudicatario='indeterminado'`, `cargo_al_credito=False`, `monto_adjudicacion=$97,993,864`

## C. Resumen

- Causas con 'cargo' en log_decision: **6**
- Con PDF disponible y re-analizadas: **6**
- Sin PDF (no re-analizables): **0**
- **SOSPECHOSAS de eliminación errónea (tercero + monto >= deuda): 0**