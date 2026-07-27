# CIERRE DE PROYECTO — REGIONES (Remates Judiciales, fuente PyL)

**Estado: PAUSA INDEFINIDA**
**Fecha de cierre: 27 de julio de 2026**
**Ultimo commit funcional validado: a86bd1d (Tanda E, 12 de junio de 2026)**
**Decision tomada por: Diego (orquestador)**

---

## 1. Por que se pausa

La Oficina Judicial Virtual del Poder Judicial esta protegida por F5 BIG-IP
con Shape Security (Distributed Cloud Bot Defense), una defensa anti-bot
comercial de nivel bancario. Durante 2026 su despliegue se endurecio: el
bloqueo dejo de ser anecdotico y paso a ser frecuente.

Precision importante, para no cerrar con un motivo equivocado: el WAF **no
es invencible**. En el proyecto hermano Matriz Causas Activas se documento
una receta que logro procesar 10 causas de punta a punta sin un solo
bloqueo (perfil limpio, tecleo humano digito a digito, reset de navegador
entre causas y pausa de enfriamiento de 3 minutos). Ver
MEMORIA_TECNICA_WAF_PJUD_2026-07-05.md.

El problema es que esa receta es **serial y lenta**, y Regiones es
**paralelo y masivo**: lotes de ~67 causas con 5 workers simultaneos en
unos 7 minutos. El paralelismo es justamente la señal que enciende al WAF.
Portar la receta obliga a reescribir M2 a ejecucion serial y a aceptar
lotes de varias horas.

Conclusion honesta: no se cierra porque sea imposible, sino porque el costo
de seguir es incompatible con el volumen y la arquitectura del proyecto, y
porque la carrera armamentistica contra un producto comercial dedicado no
tiene final.

---

## 2. Que quedo funcionando (no depende del PJUD)

- **M1 v2** (v2_experimental/modulo1_v2.py): parsing de los DOCX mensuales
  de Diarios Publicos y Legales.
- **modulo1_parser.py**: biblioteca compartida de v2 y del harness.
- **filtro_cbr.py**: filtro de antiguedad CBR, en el repo y validado.
  Metrica de oro = 0 en 106 de 106 bloques auditados. Detalle en
  CIERRE_CICLO_REGIONES_2026-06.md.
- **Generacion de DOCX de auditoria para abogados**: formato v2 (Audit3),
  documentado en FORMATO_DOCX_AUDITORIA.md.
- **Pipeline docx -> Excel** (Regiones_docxToExcel.bat).

---

## 3. Que dejo de funcionar (la causa del cierre)

- **M2** (modulo2_ojv.py / ojv_remates.py): verificacion de causas en la
  OJV. Bloqueado por el WAF en su modo paralelo de 5 workers.
- **filtrador_saldos.py** (Detector_Excedentes.bat): deteccion de causas
  con saldo. Depende del mismo scraping. Bloqueado.

---

## 4. Que quedo a medias (con ubicacion exacta, para no repetir el diagnostico)

- **Tanda C1 — admitir "/" como separador de anio en el ROL.** Diseño
  cerrado, ACT escrito, NUNCA aplicado (verificado: modulo1_v2.py sin
  modificar desde el 10-jun). Son tres puntos: _RE_TIENE_ROL en
  v2_experimental/modulo1_v2.py; _RE_ROL y _RE_ROL_HEADER en
  modulo1_parser.py; y la replica de _RE_TIENE_ROL en test_cbr_docx.py.
  El cambio es unico: añadir "/" dentro de la clase de separadores
  [-–—]. Predicciones de validacion ya fijadas: pre-filtro sin ROL
  12 -> 11, con ROL 94 -> 95, causas nuevas 0 -> 1 (rol 5702-2024).
- **Tanda C2 — cohorte "Rol N°" sin letra.** Fase 1 (medicion offline)
  nunca corrida. Requiere reponer corpus historico en Diarios\, que hoy
  contiene solo el DOCX del 18 al 22 de mayo de 2026.
- **Vigilancia mensual del filtro CBR (P3).** Sin insumo: no llego
  material nuevo de PyL despues de mayo.
- **Anotados de calidad (P4).** Pendiente menor de documentacion.
- **Frente de saldos (P5).** Congelado en el commit "WIP congelado" de
  este cierre. ADVERTENCIA: ese codigo NO fue validado ni corrido de punta
  a punta. No asumir que funciona.
- **Pendiente de negocio (sin codigo):** nunca se confirmo con el abogado
  si el CSV de descartes del pre-filtro sirve como cola de revision manual.

---

## 5. La puerta que queda abierta

Si algun dia se quiere extraer valor sin tocar el PJUD: M1 + filtro CBR +
generacion de DOCX de auditoria funcionan de forma autonoma sobre los DOCX
de PyL. Lo que se pierde es la verificacion en la OJV y la deteccion de
saldos; es decir, el embudo 270 -> 22 -> 6 pasaria a depender por completo
del ojo del abogado. Es un producto mas pobre, pero vivo y sin adversario.

---

## 6. Condicion de reactivacion (no reactivar por impulso)

Cualquiera de estas señales objetivas:
- a) El PJUD habilita una via de consulta programatica legitima.
- b) Se comprueba empiricamente que una sesion autenticada de abogado
  recibe trato mas laxo del WAF. Hipotesis abierta y NO comprobada; ver
  seccion 5.3 de MEMORIA_TECNICA_WAF_PJUD_2026-07-05.md.
- c) Se decide aceptar lotes seriales de varias horas y reescribir M2 con
  la receta del proyecto hermano.

---

## 7. Donde esta todo

- Repositorio: D:\Remates -> github.com/DFNR2026/RemateRegiones (main
  sincronizado al cerrar).
- Reglas de trabajo del ejecutor: .clinerules
- Memoria del adversario: MEMORIA_TECNICA_WAF_PJUD_2026-07-05.md,
  AVISO_WAF_PJUD_para_otros_proyectos.md e INCIDENTE_WAF_2026-04-09.md
- Estado del ultimo ciclo tecnico: CIERRE_CICLO_REGIONES_2026-06.md
- Ultimo traspaso operativo: TRASPASO_PENDIENTES_2026-06-12.md
- Aprendizajes de la auditoria con abogados: RESUMEN_CIERRE_AUDIT*.md

---

## 8. Advertencia para el yo del futuro

Los proyectos hermanos (Mercurio/RM y Matriz Causas Activas) TAMBIEN
consultan la OJV. Cerrar Regiones no elimina la dependencia del PJUD del
ecosistema completo: solo retira de la mesa el consumidor mas voluminoso.
Que hacen los hermanos con esa dependencia es una decision aparte, no
tomada en este cierre.

---

*Fin del documento de cierre.*