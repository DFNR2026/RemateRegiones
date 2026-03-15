# ============================================================
# AUTOMATIZACIÓN OJV - DESCARGA DOCUMENTOS DE REMATES
# Versión 10.2
#
# INPUT:  D:\Remates\causas_ojv.xlsx  (hoja CAUSAS)
# OUTPUT: D:\Remates\Descargas\C-ROL-AÑO_MANDAMIENTO.pdf   (solo Ejecutivos)
#                               C-ROL-AÑO_BASES_REMATE.pdf  (solo Ley de Bancos)
#
# Lógica:
#   - Ejecutivo Obl. de Dar → solo Mandamiento (dirección ya viene del diario)
#   - Ley de Bancos          → solo Bases de Remate (contienen mínimo subasta)
#
# Uso:    python D:\Remates\ojv_remates.py
# Detener: Ctrl+C
#
# ===========================================================================
# ANTES DE CADA CORRIDA DE PRUEBA — limpiar estado previo:
#
#   python limpiar_cache.py
#
# Esto borra:
#   1. Todo el contenido de D:\\Remates\\Descargas\\
#   2. Las filas de datos de la hoja CAUSAS en causas_ojv.xlsx
#      (deja solo el encabezado, para que M1 no filtre causas ya procesadas)
#
# Por qué es importante:
#   - Sin limpiar Descargas: los PDFs de la corrida anterior permanecen.
#     Si un archivo ya existe con el nombre correcto pero contenido de otra
#     causa (bug de "primera fila"), el error pasa desapercibido.
#   - Sin limpiar CAUSAS: M1 ve todas las causas como "ya procesadas"
#     y devuelve 0 causas → el pipeline termina inmediatamente.
# ===========================================================================
# ============================================================

from playwright.sync_api import sync_playwright
import pandas as pd
import os, re, time, unicodedata
from rapidfuzz import fuzz

EXCEL_INPUT       = r"D:\Remates\causas_ojv.xlsx"
CARPETA_DESCARGAS = r"D:\Remates\Descargas"


# ============================================================
# LECTURA DEL EXCEL  (columnas: ROL, AÑO, CORTE, TRIBUNAL)
# ============================================================
def leer_causas():
    try:
        df = pd.read_excel(EXCEL_INPUT, sheet_name="CAUSAS", dtype=str)
        df = df.dropna(subset=["ROL", "AÑO", "CORTE", "TRIBUNAL"])
        df = df[df["ROL"].str.strip().str.match(r"^\d+$")]
        return [
            {
                "rol":      r["ROL"].strip(),
                "año":      r["AÑO"].strip(),
                "corte":    r["CORTE"].strip(),
                "tribunal": r["TRIBUNAL"].strip(),
            }
            for _, r in df.iterrows()
        ]
    except FileNotFoundError:
        print(f"\n  ✗ Excel no encontrado: {EXCEL_INPUT}")
        return []
    except Exception as e:
        print(f"\n  ✗ Error leyendo Excel: {e}")
        return []


# ============================================================
# HELPERS
# ============================================================
def cerrar_popups(page):
    for sel in ["button.close", ".modal-footer button",
                "button:has-text('Aceptar')", "button:has-text('Cerrar')"]:
        try:
            if page.is_visible(sel, timeout=500):
                page.click(sel)
                time.sleep(0.3)
        except:
            pass


def cerrar_modal_aviso(page):
    """Detecta el modal de avisos del PJ (#close-modal) y lo cierra con Escape."""
    try:
        if page.is_visible("#close-modal", timeout=800):
            page.keyboard.press("Escape")
            time.sleep(0.3)
    except Exception:
        pass


def _quitar_tildes(s: str) -> str:
    """Elimina diacríticos: 'Ángeles' → 'Angeles', 'Concepción' → 'Concepcion'."""
    return "".join(
        c for c in unicodedata.normalize("NFD", s)
        if unicodedata.category(c) != "Mn"
    )


def _normalizar_texto_ojv(texto: str) -> str:
    """Normalización agresiva para comparar tribunales contra dropdown OJV."""
    t = texto.lower()
    t = "".join(c for c in unicodedata.normalize("NFD", t) if unicodedata.category(c) != "Mn")
    t = re.sub(r'\b(primer|primera|1er)\b', '1', t)
    t = re.sub(r'\b(segundo|segunda|2do)\b', '2', t)
    t = re.sub(r'\b(tercer|tercero|tercera|3er)\b', '3', t)
    t = re.sub(r'\b(cuarto|cuarta|4to)\b', '4', t)
    t = re.sub(r'\b(quinto|quinta|5to)\b', '5', t)
    t = re.sub(r'[°º\xba\.,-]', ' ', t)
    t = re.sub(r'\bde\b', ' ', t)
    t = t.replace("jdo", "juzgado")
    t = t.replace("garantia", "gar")
    # "Juzgado de Letras Civil" y "Juzgado de Letras" son lo mismo
    # (el diario a veces agrega "Civil" pero el dropdown OJV no lo tiene)
    t = re.sub(r'\bletras\s+civil\b', 'letras', t)
    return re.sub(r'\s+', ' ', t).strip()


def seleccionar_por_texto(page, selector_id, texto_buscar, timeout_seg=15):
    """Selecciona opción en <select> usando fuzzy matching (RapidFuzz) contra el DOM."""
    norm_buscar = _normalizar_texto_ojv(texto_buscar)
    print(f"    Buscando '{texto_buscar}' en #{selector_id}...", end="", flush=True)

    mejor_score = 0
    mejor_valor = None
    mejor_texto = None

    for intento in range(timeout_seg * 2):
        try:
            opciones = page.query_selector_all(f"#{selector_id} option")

            mejor_score = 0
            mejor_valor = None
            mejor_texto = None

            for op in opciones:
                texto_opcion = op.inner_text().strip()
                if not texto_opcion or texto_opcion.lower() == "todos":
                    continue

                norm_opcion = _normalizar_texto_ojv(texto_opcion)
                score = fuzz.token_set_ratio(norm_buscar, norm_opcion)

                if score > mejor_score:
                    mejor_score = score
                    mejor_valor = op.get_attribute("value")
                    mejor_texto = texto_opcion

            if mejor_score >= 80 and mejor_valor:
                page.select_option(f"#{selector_id}", value=mejor_valor)
                print(f" ✓ '{mejor_texto}' ({mejor_score:.1f}%)")
                return True

        except Exception:
            pass

        time.sleep(0.5)
        print(".", end="", flush=True)

    print(f" ✗ no encontrado (mejor score: {mejor_score:.1f}%, umbral: 80%)")
    return False


# ============================================================
# NAVEGACIÓN
# ============================================================
def navegar_a_consulta(page):
    """Navega a indexN.php para evitar el modal de avisos de home/index.php.
    Si la OJV redirige a home/index.php, hace click en 'Consulta causas'
    para llegar al formulario (el modal no aparece en redirecciones)."""
    print("  -> Cargando OJV...")
    try:
        page.goto(
            "https://oficinajudicialvirtual.pjud.cl/indexN.php",
            wait_until="domcontentloaded",
            timeout=30000
        )
    except Exception as e:
        print(f"  ⚠ {e}")
        # goto puede expirar mientras el redirect indexN→home/index.php sigue en curso.
        # Esperar a que la navegación pendiente termine antes de revisar la URL.
        try:
            page.wait_for_load_state("domcontentloaded", timeout=15000)
        except Exception:
            pass
    try:
        if "home/index.php" in page.url:
            cerrar_modal_aviso(page)
            page.wait_for_selector("text=Consulta causas", timeout=10000)
            page.click("text=Consulta causas")
        page.wait_for_selector("#competencia", timeout=15000)
        time.sleep(0.8)
        print("  ✓ Formulario listo")
        return True
    except Exception as e:
        print(f"  ✗ No se abrió el formulario: {e}")
        return False


def limpiar_formulario(page):
    """Limpia el formulario para una nueva búsqueda sin navegar a ninguna URL.
    1. Cierra modal si está visible.
    2. Intenta botón 'Limpiar' del formulario OJV.
    3. Si no existe, resetea #competencia manualmente (AJAX limpia los dependientes)."""
    cerrar_modal_aviso(page)
    try:
        if page.is_visible("button:has-text('Limpiar')", timeout=1000):
            page.click("button:has-text('Limpiar')")
            time.sleep(0.5)
            return True
    except Exception:
        pass
    try:
        page.select_option("#competencia", value="")
        time.sleep(0.3)
        return True
    except Exception as e:
        print(f"  ⚠ Error limpiando formulario: {e}")
        return False


# ============================================================
# BÚSQUEDA  (Corte y Tribunal por TEXTO, no por ID)
# ============================================================

def buscar_causa(page, rol, año, corte, tribunal):
    """Búsqueda exacta: competencia + corte + tribunal + libro C + ROL + año."""
    print(f"  -> Buscando C-{rol}-{año}...")
    try:
        # 1. Competencia = Civil
        page.select_option("#competencia", value="3")
        time.sleep(0.8)

        # 2. Corte
        if not seleccionar_por_texto(page, "conCorte", corte, timeout_seg=10):
            print("  ✗ Corte no encontrada")
            return False
        time.sleep(1.0)

        # 3. Tribunal
        if not seleccionar_por_texto(page, "conTribunal", tribunal, timeout_seg=15):
            print("  ✗ Tribunal no encontrado")
            return False
        time.sleep(1.5)   # dar tiempo al AJAX antes de tocar ROL/año

        # 4. Libro = C
        page.select_option("#conTipoCausa", value="C")
        time.sleep(0.5)

        # 5. ROL + año — verificar que AJAX no los borró
        page.fill("#conRolCausa", "")
        page.fill("#conEraCausa", "")
        page.fill("#conRolCausa", rol)
        page.fill("#conEraCausa", año)
        time.sleep(0.3)

        if page.input_value("#conRolCausa") != rol or page.input_value("#conEraCausa") != año:
            print(f"  ⚠ Campos borrados por AJAX — reintentando fill")
            page.fill("#conRolCausa", rol)
            page.fill("#conEraCausa", año)
            time.sleep(0.5)

        # 6. Buscar
        page.click("#btnConConsulta")
        page.wait_for_load_state("domcontentloaded", timeout=20000)
        time.sleep(2.5)
        cerrar_modal_aviso(page)
        cerrar_popups(page)

        # Detectar mensaje "sin resultados" (OJV lo muestra aunque haya filas de layout)
        try:
            if "No se han encontrado resultados" in page.inner_text("body"):
                print("  ✗ Sin resultados (OJV: causa no encontrada)")
                return False
        except Exception:
            pass

        # Contar solo filas reales (filtrando tablas de layout de la página)
        filas = page.query_selector_all("table#veDetalle tbody tr")
        if not filas:
            filas = [
                f for f in page.query_selector_all("table tbody tr")
                if f.inner_text().strip()
                and "No se han encontrado" not in f.inner_text()
                and "VALOR RECUSACIÓN" not in f.inner_text()
            ]
        if not filas:
            print("  ✗ Sin resultados")
            return False
        print(f"  ✓ {len(filas)} resultado(s)")
        return True

    except Exception as e:
        print(f"  ✗ Error en búsqueda: {e}")
        return False


# ============================================================
# ABRIR DETALLE  (lupa dentro de la tabla de resultados)
# ============================================================
def abrir_detalle(page, rol=None, año=None):
    """
    Abre el modal de detalle de la causa correcta en la tabla de resultados.

    Busca la fila que contenga el ROL exacto.  Si no la encuentra, NO usa
    primera fila como fallback (evita abrir la causa equivocada).

    Hace scroll_into_view_if_needed() antes del click para evitar el error
    "element is not visible" cuando la fila queda fuera del viewport.
    """
    try:
        filas = (
            page.query_selector_all("table#veDetalle tbody tr") or
            page.query_selector_all("table tbody tr")
        )
        if not filas:
            print("  ✗ Sin filas en tabla de resultados")
            return False

        # Buscar fila que contenga el ROL buscado
        fila_objetivo = None
        if rol and año:
            año2d = año[-2:]   # "2023" → "23"  (OJV puede mostrar año en 2 dígitos)
            for fila in filas:
                texto = fila.inner_text()
                # Acepta año en 4 dígitos ("2023") o 2 dígitos ("23")
                año_ok = año in texto or año2d in texto
                if rol in texto and año_ok:
                    fila_objetivo = fila
                    break

            if fila_objetivo is None:
                # Diagnóstico: mostrar primeras 3 filas para entender el formato
                print(f"  ✗ ROL {rol}-{año} no encontrado en {len(filas)} filas")
                for i, f in enumerate(filas[:3]):
                    print(f"    fila[{i}]: {f.inner_text()[:100].strip()!r}")
                return False

        if fila_objetivo is None:
            fila_objetivo = filas[0]

        # Localizar la lupa dentro de la fila objetivo
        lupa = (
            fila_objetivo.query_selector("a.toggle-modal") or
            fila_objetivo.query_selector("td:first-child a") or
            fila_objetivo.query_selector("a[href='#modalDetalleCivil']") or
            fila_objetivo.query_selector("a")
        )
        if not lupa:
            print("  ✗ Lupa no encontrada en la fila")
            return False

        # Si encontramos el <i>, subir al <a> padre
        tag = lupa.evaluate("el => el.tagName.toLowerCase()")
        if tag == "i":
            lupa = lupa.evaluate_handle("el => el.closest('a') || el.parentElement")

        lupa.scroll_into_view_if_needed()
        lupa.click()
        time.sleep(2.5)

        try:
            page.wait_for_selector("#modalDetalleCivil, .modal.in, .modal.show", timeout=7000)
        except:
            pass

        print("  ✓ Detalle abierto")
        return True

    except Exception as e:
        print(f"  ✗ Error abriendo detalle: {e}")
        return False


# ============================================================
# CUADERNO  (select#selCuaderno — por texto)
# ============================================================
def seleccionar_cuaderno(page, texto):
    try:
        page.wait_for_selector("#selCuaderno", timeout=8000)
        return seleccionar_por_texto(page, "selCuaderno", texto, timeout_seg=8)
    except Exception as e:
        print(f"  ✗ Error cuaderno: {e}")
        return False


def filas_del_modal(page):
    """Devuelve filas de la tabla DENTRO del modal (no de la tabla de resultados)."""
    selectores = [
        "#loadHistCuadernoCivil table tbody tr",
        "#historiaClv table tbody tr",
        "#modalDetalleCivil .table-responsive table tbody tr",
        "#modalDetalleCivil table tbody tr",
        ".modal.in table tbody tr",
        ".modal.show table tbody tr",
    ]
    for sel in selectores:
        try:
            page.wait_for_selector(sel, timeout=1500)
            filas = page.query_selector_all(sel)
            if filas:
                print(f"    Tabla modal: {sel} ({len(filas)} filas)")
                return filas
        except:
            continue
    print("    ⚠ Tabla del modal no encontrada")
    return []


def descargar_pdf_de_fila(page, context, fila, nombre_archivo):
    """
    HTML real del inspector:
      <form action="/civil/documentos/doculs.php" method="get" target="p3">
        <input type="hidden" name="dtaDoc" value="JWT">
        <a href="#" onclick=".closest(form).submit();" title="Descargar Documento">
          <i class="fa fa-file-pdf-o fa-lg"></i>
        </a>
      </form>
    """
    enlace = (
        fila.query_selector("a[title='Descargar Documento']") or
        fila.query_selector("form a") or
        fila.query_selector("a[onclick*='submit']") or
        fila.query_selector("i.fa-file-pdf-o")
    )
    if not enlace:
        print("    ✗ Botón descarga no encontrado")
        return False

    if enlace.evaluate("el => el.tagName.toLowerCase()") == "i":
        enlace = enlace.evaluate_handle("el => el.closest('a') || el")

    # form target="p3" abre nueva pestaña
    try:
        with context.expect_page(timeout=15000) as popup_info:
            enlace.click()
        popup = popup_info.value
        popup.wait_for_load_state("domcontentloaded", timeout=15000)
        url_pdf = popup.url
        time.sleep(1)
        try:
            response = popup.request.get(url_pdf, timeout=20000)
            if response.ok and len(response.body()) > 500:
                with open(nombre_archivo, "wb") as f:
                    f.write(response.body())
                print(f"    ✓ {os.path.basename(nombre_archivo)}")
                popup.close()
                return True
        except:
            pass
        popup.close()
    except:
        pass

    try:
        with page.expect_download(timeout=20000) as dl:
            enlace.click()
        dl.value.save_as(nombre_archivo)
        print(f"    ✓ {os.path.basename(nombre_archivo)}")
        return True
    except Exception as e:
        print(f"    ✗ Descarga fallida: {e}")
        return False


def buscar_mandamiento(page, context, etiqueta):
    nombre = os.path.join(CARPETA_DESCARGAS, f"{etiqueta}_MANDAMIENTO.pdf")
    filas  = filas_del_modal(page)
    if not filas:
        return False

    tiene_pdf = "a[title='Descargar Documento'], form a, i.fa-file-pdf-o"

    # OPTIMIZACIÓN: La tabla viene ordenada por folio DESCENDENTE (ej: 68, 67, ...1)
    # El Mandamiento suele estar entre folio 1 y 5, así que buscamos desde el FINAL
    filas_rev = list(reversed(filas))

    # Diagnóstico: mostrar últimas filas (folios bajos, donde está el mandamiento)
    for i, fila in enumerate(filas_rev[:5]):
        celdas = fila.query_selector_all("td")
        txts = [c.inner_text().strip() for c in celdas]
        print(f"    fila[folio bajo {i}]: {txts}")

    # Estructura real de columnas:
    # txts[0]=Folio | txts[1]=Doc(sr-only) | txts[2]=Anexo
    # txts[3]=Etapa | txts[4]=Tramite | txts[5]=Desc.Tramite
    #
    # El Mandamiento (folio 1) tiene:
    #   Etapa="Mandamiento", Tramite="" (vacío), Desc="Mandamiento"
    # => buscar Etapa(3)="mandamiento" + Desc(5)="mandamiento"

    for fila in filas_rev:
        celdas = fila.query_selector_all("td")
        txts = [c.inner_text().strip().lower() for c in celdas]
        if len(txts) >= 6:
            etapa = txts[3]
            desc  = txts[5]
            if etapa == "mandamiento" and desc == "mandamiento":
                if fila.query_selector(tiene_pdf):
                    print(f"    ✓ Mandamiento exacto folio {txts[0]}")
                    return descargar_pdf_de_fila(page, context, fila, nombre)

    # Fallback: Etapa="mandamiento" sin requerimiento en ninguna celda
    for fila in filas_rev:
        celdas = fila.query_selector_all("td")
        txts = [c.inner_text().strip().lower() for c in celdas]
        fila_texto = " ".join(txts)
        if len(txts) >= 4 and txts[3] == "mandamiento" and "requerimiento" not in fila_texto:
            if fila.query_selector(tiene_pdf):
                print(f"    ✓ Mandamiento (Etapa) folio {txts[0]}")
                return descargar_pdf_de_fila(page, context, fila, nombre)

    print("    ⚠ Mandamiento no encontrado — revisa filas mostradas arriba")
    return False


def buscar_bases_remate(page, context, etiqueta):
    nombre     = os.path.join(CARPETA_DESCARGAS, f"{etiqueta}_BASES_REMATE.pdf")
    filas      = filas_del_modal(page)
    candidatas = []

    tiene_pdf = "a[title='Descargar Documento'], form a, i.fa-file-pdf-o"
    for fila in filas:
        texto = fila.inner_text().lower()
        # Incluir solo "propone bases" — EXCLUIR "aprueba"/"aprobada" (resolución del juez)
        es_propone = "propone bases" in texto or "bases de remate" in texto
        es_aprobada = "aprueba" in texto or "aprobada" in texto or "aprobado" in texto
        if es_propone and not es_aprobada and fila.query_selector(tiene_pdf):
            candidatas.append(fila)

    if not candidatas:
        print("    ⚠ 'Propone bases de remate' no encontrado")
        return False

    # Si hay varias, usar la ÚLTIMA (folio más alto = más reciente = anterior rechazada)
    fila_objetivo = candidatas[-1]
    if len(candidatas) > 1:
        print(f"    ⚠ {len(candidatas)} versiones de bases — usando la más reciente (folio {fila_objetivo.query_selector_all('td')[0].inner_text().strip()})")
    else:
        print(f"    ✓ Bases propuestas encontradas")
    return descargar_pdf_de_fila(page, context, fila_objetivo, nombre)


# ============================================================
# EXTRACCIÓN DE LITIGANTES
# ============================================================
def _extraer_litigantes(page, rol: str) -> dict:
    """
    Extrae nombres de DTE (demandante) y DDO (demandado) desde la
    pestaña Litigantes de la OJV.

    Returns: {'nombre_dte': str|None, 'nombre_ddo': str|None}
    """
    resultado = {'nombre_dte': None, 'nombre_ddo': None}

    try:
        # 1. Click en la pestaña Litigantes (puede que ya esté activa)
        tab_link = page.locator('a[href="#litigantesCiv"]')
        if tab_link.count() > 0:
            tab_link.click()
            page.wait_for_selector(
                '#litigantesCiv tbody tr',
                state='visible',
                timeout=8000
            )
        else:
            print(f"  [LITIGANTES] {rol}: pestaña no encontrada")
            return resultado

        # 2. Leer SOLO la tabla dentro de #litigantesCiv
        filas = page.locator('#litigantesCiv tbody tr').all()

        for fila in filas:
            celdas = fila.locator('td').all()
            if len(celdas) < 4:
                continue

            participante = celdas[0].inner_text().strip()
            nombre_raw = celdas[3].inner_text().strip()

            # Limpiar nombre: quitar "(Poder Amplio)" y trailing spaces
            nombre = re.sub(r'\s*\(.*?\)\s*$', '', nombre_raw).strip()
            nombre = re.sub(r'\s+', ' ', nombre)

            if participante.startswith('DDO'):
                if resultado['nombre_ddo'] is None:  # tomar solo el primero
                    resultado['nombre_ddo'] = nombre
            elif participante.startswith('DTE'):
                if resultado['nombre_dte'] is None:  # tomar solo el primero
                    resultado['nombre_dte'] = nombre

        # Log de resultados
        if resultado['nombre_dte'] or resultado['nombre_ddo']:
            print(f"  ✓ Litigantes {rol}: DTE={resultado['nombre_dte']} | DDO={resultado['nombre_ddo']}")
        else:
            print(f"  [LITIGANTES] {rol}: tabla visible pero sin DTE/DDO")

    except Exception as e:
        print(f"  [LITIGANTES] {rol}: {e}")

    return resultado


# ============================================================
# PROCESAMIENTO
# ============================================================
def procesar_causa(page, context, causa):
    etiqueta = f"C-{causa['rol']}-{causa['año']}"
    print(f"\n{'='*55}")
    print(f"  Causa    : {etiqueta}")
    print(f"  Corte    : {causa['corte']}")
    print(f"  Tribunal : {causa['tribunal']}")
    print(f"{'='*55}")

    if not buscar_causa(page, causa["rol"], causa["año"],
                        causa["corte"], causa["tribunal"]):
        return False, False
    if not abrir_detalle(page, causa["rol"], causa["año"]):
        return False, False

    # ── Extraer litigantes desde pestaña #litigantesCiv ─────
    litigantes = _extraer_litigantes(page, etiqueta)
    if litigantes['nombre_dte']:
        causa['demandante'] = litigantes['nombre_dte']
    if litigantes['nombre_ddo']:
        causa['demandado'] = litigantes['nombre_ddo']

    # ── Detectar tipo de procedimiento ──────────────────────
    # "Proc.: Ejecutivo Obligación de Dar"  → cuaderno Apremio
    # "Proc.: Ley de Bancos"                → cuaderno Principal
    es_ley_bancos = False
    try:
        modal = page.query_selector("#modalDetalleCivil, .modal.in, .modal.show")
        if modal:
            modal_texto = modal.inner_text().lower()
            es_ley_bancos = "ley de bancos" in modal_texto
    except:
        pass

    cuaderno_objetivo = "Principal" if es_ley_bancos else "Apremio"
    tipo_proc = "Ley de Bancos -> cuaderno Principal" if es_ley_bancos else "Ejecutivo -> cuaderno Apremio"
    print(f"  Proc.: {tipo_proc}")

    mandamiento_ok = bases_ok = False

    if seleccionar_cuaderno(page, cuaderno_objetivo):
        if es_ley_bancos:
            # Ley de Bancos: no existe Mandamiento (hipoteca ya constituida)
            # Solo descargar Bases de Remate desde cuaderno Principal
            # (contienen el mínimo de la subasta en un otrosí)
            mandamiento_ok = True   # marcar como N/A (no aplica)
            print("\n  [MANDAMIENTO] N/A — Ley de Bancos (no aplica)")
            print("\n  [BASES DE REMATE]")
            bases_ok = buscar_bases_remate(page, context, etiqueta)
        else:
            # Ejecutivo Obligación de Dar: solo Mandamiento desde cuaderno Apremio
            # (la dirección del inmueble ya viene del aviso del diario P&L,
            #  por lo que las bases de remate ya no son necesarias)
            bases_ok = True   # marcar como N/A (no aplica)
            print("\n  [MANDAMIENTO]")
            mandamiento_ok = buscar_mandamiento(page, context, etiqueta)
            print("\n  [BASES DE REMATE] N/A — Ejecutivo (dirección viene del diario)")
    else:
        print(f"  ✗ Cuaderno '{cuaderno_objetivo}' no disponible")

    # Cerrar modal
    for sel in ["button:has-text('Cerrar')", ".modal .close", "button.close"]:
        try:
            page.click(sel, timeout=2000); break
        except: pass
    try: page.keyboard.press("Escape")
    except: pass
    time.sleep(1)

    estado_mand = "N/A (Ley Bancos)" if es_ley_bancos else ("✓" if mandamiento_ok else "✗ FALTA")
    estado_bases = "N/A (Ejecutivo)" if not es_ley_bancos else ("✓" if bases_ok else "✗ FALTA")
    print(f"\n  Resultado {etiqueta}:")
    print(f"    Mandamiento  : {estado_mand}")
    print(f"    Bases remate : {estado_bases}")
    return mandamiento_ok, bases_ok


# ============================================================
# MAIN
# ============================================================
def main():
    print("\n" + "="*55)
    print("  AUTOMATIZACIÓN OJV — REMATES JUDICIALES CHILE")
    print("  Versión 10.0  (Ejecutivo=Mandamiento | LeyBancos=Bases)")
    print("="*55)
    print(f"  Excel : {EXCEL_INPUT}")
    print(f"  Salida: {CARPETA_DESCARGAS}")
    print("  ▶ Detener: Ctrl+C")
    print("="*55)

    causas = leer_causas()
    if not causas:
        print("\n  Sin causas. Revisa el Excel.")
        return

    print(f"\n  {len(causas)} causa(s):")
    for c in causas:
        print(f"    C-{c['rol']}-{c['año']}  |  {c['corte']}  |  {c['tribunal']}")

    os.makedirs(CARPETA_DESCARGAS, exist_ok=True)
    ok, parcial, error = [], [], []

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False, slow_mo=100)
        context = browser.new_context(accept_downloads=True)
        page    = context.new_page()
        page.set_default_timeout(15000)

        if not navegar_a_consulta(page):
            print("  ✗ No se pudo abrir el formulario OJV")
            browser.close()
            return

        for i, causa in enumerate(causas, 1):
            etiqueta = f"C-{causa['rol']}-{causa['año']}"
            print(f"\n[{i}/{len(causas)}] {etiqueta}")
            if i > 1:
                limpiar_formulario(page)
            try:
                m, b = procesar_causa(page, context, causa)
                if m and b:   ok.append(etiqueta)
                elif m or b:  parcial.append(f"{etiqueta} ({'M' if m else '-'}/{'B' if b else '-'})")
                else:         error.append(etiqueta)
            except KeyboardInterrupt:
                print("\n\n  ⏸ Detenido")
                break
            except Exception as e:
                print(f"  ✗ ERROR: {e}")
                error.append(etiqueta)
            time.sleep(2)

        browser.close()

    print("\n" + "="*55)
    print("  RESUMEN")
    print("="*55)
    print(f"\n  ✓ Completos  ({len(ok)}):"); [print(f"    {c}") for c in ok]
    print(f"\n  ⚠ Parciales  ({len(parcial)}) M=Mandamiento B=Bases:"); [print(f"    {c}") for c in parcial]
    print(f"\n  ✗ Sin descarga ({len(error)}):"); [print(f"    {c}") for c in error]
    print(f"\n  Archivos en: {CARPETA_DESCARGAS}")
    print(f"\n  Nota: Ejecutivos solo descargan Mandamiento")
    print(f"        Ley de Bancos solo descarga Bases de Remate")
    print("="*55)


if __name__ == "__main__":
    main()
