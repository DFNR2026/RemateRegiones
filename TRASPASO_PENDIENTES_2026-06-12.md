# TRASPASO — Pendientes de Regiones (generado desde el chat del proyecto Mercurio)
**Fecha:** 12 de junio de 2026
**Origen:** chat de mentoría/auditoría en el proyecto "El Mercurio" (ciclo cerrado en `CIERRE_CICLO_REGIONES_2026-06.md`).
**Destino:** nuevo chat en el proyecto "Regiones".

---

## 0. PRIMERA INSTRUCCIÓN PARA EL NUEVO CHAT (antes de ejecutar nada)

Claude: usa la búsqueda de conversaciones pasadas **de este proyecto (Regiones)** y
localiza el último chat (hace ~1 semana, Diego no lo abre desde entonces). Revisa
qué se estaba discutiendo y en qué quedó. Concilia contra este documento:
- Si algo de allá ya quedó resuelto por las Tandas A/B1/B2/D de este ciclo → márcalo cerrado.
- Si algo sigue abierto y NO está en la lista de abajo → incorpóralo a la planificación
  con su prioridad, y preséntale a Diego el plan unificado antes del primer ACT.

---

## 1. ESTADO ACTUAL (verificado a nivel de refs en disco)

- Producción: **v2** (`v2_experimental/modulo1_v2.py`, Regex+Haiku). v1 **eliminado** (Tanda D).
- `modulo1_parser.py` = biblioteca compartida oficial (~650 líneas; docstring lo declara).
- Red de seguridad del pre-filtro activa: `logs/descartes_prefiltro_v2_*.csv`.
- Pre-filtro = superconjunto del extractor (fix punto de miles, B2).
- Veredicto CBR: rama ambigua NO muerde (106/106, métrica de oro 0). Fix asimétrico archivado.
- `local main` = `origin/main` = `fbecfa1`. Hashes canónicos del ciclo (post-rebase):
  `aaa8723` (D 1/3), `6efec23` (D 2/3), `3b122af` (D 3/3), `fbecfa1` (filtro_cbr al repo).
  Los hashes pre-rebase que aparecen en `MD_Cline/` están obsoletos.
- Working tree: 4 archivos modificados del **frente de saldos** (filtrador_saldos.py,
  modulo5_reporte.py [línea "CBR Motivo"], Causas_posible_saldo.xlsx,
  causas_eliminadas_historial.csv). NO tocarlos desde este frente.

## 2. DECRETOS DE NEGOCIO VIGENTES (Diego, 12-jun)

- **A- (arbitrales) y V- (voluntarias): FUERA DE ALCANCE.** No se extraen. El diseño
  actual (prefiltro las reconoce, extraer_rol solo-C las descarta) es correcto.
- **Slash (`Rol C-5702/2024`): SE INCLUYE.** Sin más consulta. → Tanda C1.
- **`Rol Nº` sin letra: SE ANALIZA.** Heurística de Diego: clasificar por carátula
  (entidad bancaria/financiera como demandante → civil, incluir; "APELLIDO1/APELLIDO2"
  sin entidad → probable arbitral/partición, descartar) + verificación final en OJV por
  campo **Proc.** → Tanda C2.
  **Guarda obligatoria (riesgo de homónimo):** los ROL son únicos por tribunal Y tipo.
  Un sin-letra arbitral buscado en competencia civil puede devolver un homónimo civil
  ajeno. Para la cohorte sin-letra, M2 debe cruzar la carátula OJV contra
  demandante/demandado del aviso antes de aceptar el match.

## 3. PLAN PRIORIZADO

### P0 — Tanda E: .gitignore definitivo (ACT listo, dispara directo)
El .gitignore existente ya cubre mucho (logs/**, perfiles chrome-w*, backups, audit_html,
Descargas, Diarios). Faltan exactamente estos patrones (auditado contra `git status` real):

```
Modo: ACT
Workspace: D:\Remates
Acción: Tanda E — completar .gitignore. ANEXAR al final del archivo
.gitignore existente (NO reescribirlo) el bloque:

# === Tanda E (2026-06): residuos de runtime y flujo Claude/Cline ===
Causas con liq/liquidaciones/
Causas con liq/liquidaciones_raw/
Reportes/
MD_Cline/
*.bak_*
dry_run*
dryrun*
log_dryrun*
auditoria_cbr_*.csv
.chrome-profile-*/

Luego: git status --short ANTES y DESPUÉS (reporta ambos; el "después"
no debe listar liquidaciones, Reportes, MD_Cline, *.bak_*, dryruns ni
auditoria_cbr). git add .gitignore → commit
"Tanda E: gitignore para residuos de runtime (liquidaciones, reportes,
dryruns, MD_Cline)" → 1 file changed o DETENTE → push y reporta salida
verbatim. PROHIBIDO: git pull/rebase; si el push es rechazado, DETENTE.
Reporte en D:\Remates\MD_Cline\TandaE.md
```
Quedarán visibles a propósito (decisión consciente, no basura): los .md/py de
diagnóstico del frente de saldos, clinerules, .clineignore, CLAUDE_CODE_PROMPT.

### P1 — Tanda C1: slash como separador (ACT listo)
```
Modo: ACT
Workspace: D:\Remates
Acción: admitir "/" como separador antes del año en TRES regex,
solo en la clase del separador final:
1) v2_experimental/modulo1_v2.py · _RE_TIENE_ROL:
   ...\d[\d.\s]*[-–—]\s*\d{4}  →  ...\d[\d.\s]*[-–—/]\s*\d{4}
2) modulo1_parser.py · _RE_ROL y _RE_ROL_HEADER:
   \s*[-–—]\s*(\d{4})  →  \s*[-–—/]\s*(\d{4})   (en ambos)
3) test_cbr_docx.py · réplica de _RE_TIENE_ROL: mismo patrón que (1).
py_compile de los 3. SIN commit: validación primero.
Reporte en D:\Remates\MD_Cline\TandaC1.md
```
**Validación C1** (corrida: `python main.py --docx "Diarios\(8) RESUMEN REG. 18 AL 22
DE MAYO DE 2026.docx" --sin-ojv`), predicciones fijadas:
- Pre-filtro sin ROL: 12 → **11** (CSV de descartes con 11 filas; sale el aviso MAIAL).
- Con ROL: 94 → **95**.
- CAUSAS NUEVAS: 0 → **1** (rol **5702-2024**, el caso CBR-MANTENER dominio 1988).
- Resto idéntico (BancoEstado 7, año 1, Haiku ±1 si la dirección lo requiere).
Tres aciertos → commit (los 3 archivos juntos: el regex viaja con su réplica) → push.

### P2 — Tanda C2: `Rol Nº` sin letra (medición offline ANTES de tocar producción)
1) Script de medición (solo lectura, cero API) sobre TODOS los DOCX de `Diarios\`:
   regex candidato `[Rr]ol\s*(?:N[º°.]|Nro\.?)?\s*\d[\d.\s]*[-–—/]\s*\d{4}` con guarda
   negativa anti-avalúo (no precedido a ≤20 chars por "aval[úu]o" ni "contribucion").
   Volcar CSV: archivo | fragmento | carátula detectada | clasificación heurística
   (banco→INCLUIR / apellidos→DESCARTAR / dudoso→REVISAR) | ¿matchea avalúo?
2) Con los números a la vista, decidir umbral e implementar: prefiltro + extraer_rol
   extendidos para la cohorte sin-letra + clasificador de carátula + guarda de homónimo
   en M2 (sección 2). Esta tanda es la grande del próximo ciclo.

### P3 — Vigilancia mensual CBR (rutina, 1 minuto)
Por cada DOCX nuevo: `python test_cbr_docx.py "ruta.docx"` y leer UNA línea:
`EXCLUSIONES SOBRE AÑO AMBIGUO (solo PASA_PREFILTRO_V2=True)`. 0 durante 2-3 meses
→ cierre permanente; >0 → reabrir fix asimétrico archivado.

### P4 — Anotados sin accionar (calidad, baja prioridad)
- Etiqueta `Sin ROL post-parse` del resumen v2 es cajón de sastre (historial+RM+CBR+
  partidor). Renombrar/desglosar por motivo.
- `Haiku dir desactivado` deja direcciones en blanco con contador `Sin dirección: 0`.
  Confirmar si es intencional.

### P5 — Frente de saldos (NO es de este frente; solo conciencia)
Los 4 archivos modificados del working tree esperan su propio cierre y commit.

## 4. MÉTODO Y REGLAS (recordatorio para el nuevo chat)

- Español latinoamericano neutro estricto. PLAN antes de ACT; un cambio/un foco;
  validación con predicciones numéricas fijadas ANTES de correr; Claude audita diff y
  log antes de aprobar; verbatim cuando el relato no calza (10 desvíos de Cline este
  ciclo lo justifican). Cline ejecuta TODO terminal/git; Diego solo dice "listo, lee".
- Git: status --short antes de commitear; add explícito archivo por archivo; PROHIBIDO
  -am/add ./-A, reset --hard, clean, y TODO pull/rebase fuera de guion (push rechazado
  = detenerse y reportar). Un commit = una historia.
- Flujo: Cline escribe cada reporte en `D:\Remates\MD_Cline\<Tanda>.md`; Claude lo lee
  del disco vía Filesystem (conector activo a D:\Remates y D:\Mercurio) y verifica de
  primera mano lo crítico (refs de git incluidas).
