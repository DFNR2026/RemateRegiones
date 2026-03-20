"""
Test comparativo: v1 (Sonnet full) vs v2 (Regex-Flow + Haiku)

Corre ambas versiones sobre el mismo DOCX sandbox y compara campo por campo.
Guarda output comparativo en v2_experimental/logs/.

Uso:
  python v2_experimental/test_comparar.py --docx "D:/Remates/SANDBOX_V2_EDGE_CASES.docx"
  python v2_experimental/test_comparar.py --docx "D:/Remates/SANDBOX_V2_EDGE_CASES.docx" --solo-v2
"""

import argparse
import io
import os
import sys
import time
from datetime import datetime

# Agregar directorio padre al path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from modulo1_parser import parsear_docx_semanal
from v2_experimental.modulo1_v2 import parsear_docx_v2, get_api_calls, get_logs_dir


# Campos a comparar (los criticos para el pipeline)
CAMPOS_COMPARAR = ["rol", "año", "corte", "tribunal", "demandante", "direccion", "comuna"]

# Campos solo informativos (diferencias esperadas, no se cuentan como error)
CAMPOS_INFO = ["demandado", "fecha_remate", "minimo"]


class TeeWriter:
    """Escribe simultáneamente a consola y a archivo."""
    def __init__(self, file_path: str):
        self._file = open(file_path, "w", encoding="utf-8")
        self._stdout = sys.stdout

    def write(self, text):
        self._stdout.write(text)
        self._file.write(text)

    def flush(self):
        self._stdout.flush()
        self._file.flush()

    def close(self):
        self._file.close()


def _normalizar_campo(valor: str) -> str:
    """Normaliza un valor para comparacion justa."""
    if not valor:
        return ""
    return " ".join(valor.lower().split())


def comparar(v1_causas: list[dict], v2_causas: list[dict]) -> None:
    """Compara resultados v1 vs v2 campo por campo."""

    # Indexar por ROL-AÑO
    idx_v1 = {f"{c['rol']}-{c['año']}": c for c in v1_causas}
    idx_v2 = {f"{c['rol']}-{c['año']}": c for c in v2_causas}

    todas_claves = sorted(set(idx_v1.keys()) | set(idx_v2.keys()))

    print("\n" + "=" * 90)
    print("COMPARACION v1 (Sonnet) vs v2 (Regex+Haiku)")
    print("=" * 90)

    # Causas solo en una version
    solo_v1 = set(idx_v1.keys()) - set(idx_v2.keys())
    solo_v2 = set(idx_v2.keys()) - set(idx_v1.keys())

    if solo_v1:
        print(f"\n--- Solo en v1 ({len(solo_v1)}): {sorted(solo_v1)}")
    if solo_v2:
        print(f"\n--- Solo en v2 ({len(solo_v2)}): {sorted(solo_v2)}")

    # Comparar campo por campo
    comunes = sorted(set(idx_v1.keys()) & set(idx_v2.keys()))
    total_campos = 0
    campos_iguales = 0
    campos_diferentes = 0
    diferencias = []

    for clave in comunes:
        c1 = idx_v1[clave]
        c2 = idx_v2[clave]
        for campo in CAMPOS_COMPARAR:
            v1_val = str(c1.get(campo, ""))
            v2_val = str(c2.get(campo, ""))
            total_campos += 1
            if _normalizar_campo(v1_val) == _normalizar_campo(v2_val):
                campos_iguales += 1
            else:
                campos_diferentes += 1
                diferencias.append((clave, campo, v1_val, v2_val))

    # Tabla de diferencias
    if diferencias:
        print(f"\n--- Diferencias ({len(diferencias)} campos):")
        print(f"  {'ROL':<18} {'Campo':<12} {'v1 (Sonnet)':<40} {'v2 (Regex+Haiku)'}")
        print(f"  {'-'*17} {'-'*11} {'-'*39} {'-'*39}")
        for clave, campo, v1_val, v2_val in diferencias:
            v1_short = (v1_val[:37] + "...") if len(v1_val) > 40 else v1_val
            v2_short = (v2_val[:37] + "...") if len(v2_val) > 40 else v2_val
            print(f"  C-{clave:<15} {campo:<12} {v1_short:<40} {v2_short}")
    else:
        print("\n--- Sin diferencias en campos criticos!")

    # Resumen
    pct = (campos_iguales / total_campos * 100) if total_campos else 0
    print(f"\n--- Resumen:")
    print(f"  Causas v1:         {len(v1_causas)}")
    print(f"  Causas v2:         {len(v2_causas)}")
    print(f"  Causas comunes:    {len(comunes)}")
    print(f"  Solo en v1:        {len(solo_v1)}")
    print(f"  Solo en v2:        {len(solo_v2)}")
    print(f"  Campos comparados: {total_campos}")
    print(f"  Campos iguales:    {campos_iguales} ({pct:.1f}%)")
    print(f"  Campos diferentes: {campos_diferentes}")

    # Detalle Haiku usage
    haiku_count = sum(1 for c in v2_causas if c.get("_haiku_used"))
    print(f"\n--- Costos API:")
    print(f"  v1 llamadas Sonnet:  {len(v1_causas)} (todas las causas con ROL)")
    print(f"  v2 llamadas Haiku:   {get_api_calls()} (solo cuando regex no encontro direccion)")
    print(f"  v2 causas con Haiku: {haiku_count}/{len(v2_causas)}")

    # Detalle v2: causas sin direccion
    sin_dir_v2 = [f"C-{c['rol']}-{c['año']}" for c in v2_causas if not c.get("direccion")]
    if sin_dir_v2:
        print(f"\n  v2 sin direccion:    {sin_dir_v2}")

    print("=" * 90)


def main():
    parser = argparse.ArgumentParser(description="Comparar v1 vs v2 del parser")
    parser.add_argument("--docx", required=True, help="Ruta al DOCX de test")
    parser.add_argument("--solo-v2", action="store_true",
                        help="Solo correr v2 (sin comparacion, para testing rapido)")
    args = parser.parse_args()

    ruta = args.docx
    if not os.path.isfile(ruta):
        print(f"ERROR: archivo no encontrado: {ruta}")
        sys.exit(1)

    # Fix 7: Tee output a archivo de log
    logs_dir = get_logs_dir()
    os.makedirs(logs_dir, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_path = os.path.join(logs_dir, f"comparar_{ts}.log")
    tee = TeeWriter(log_path)
    sys.stdout = tee

    try:
        if args.solo_v2:
            print(f"\n{'='*60}")
            print(f"Corriendo solo v2 sobre: {os.path.basename(ruta)}")
            print(f"{'='*60}\n")

            t0 = time.time()
            v2_causas = parsear_docx_v2(ruta, ignorar_historial=True)
            t_v2 = time.time() - t0

            print(f"\nv2: {len(v2_causas)} causas en {t_v2:.1f}s | {get_api_calls()} llamadas Haiku")
            for c in v2_causas:
                haiku_flag = " [H]" if c.get("_haiku_used") else ""
                print(f"  C-{c['rol']}-{c['año']} | {c['corte']:<25} | "
                      f"{c['tribunal'][:35]:<35} | {c['direccion'][:40]}{haiku_flag}")
            print(f"\nLog comparativo: {log_path}")
            return

        # Correr v1
        print(f"\n{'='*60}")
        print(f"Corriendo v1 (Sonnet) sobre: {os.path.basename(ruta)}")
        print(f"{'='*60}\n")

        t0 = time.time()
        v1_causas = parsear_docx_semanal(ruta, ignorar_historial=True)
        t_v1 = time.time() - t0
        print(f"\nv1: {len(v1_causas)} causas en {t_v1:.1f}s")

        # Correr v2
        print(f"\n{'='*60}")
        print(f"Corriendo v2 (Regex+Haiku) sobre: {os.path.basename(ruta)}")
        print(f"{'='*60}\n")

        t0 = time.time()
        v2_causas = parsear_docx_v2(ruta, ignorar_historial=True)
        t_v2 = time.time() - t0
        print(f"\nv2: {len(v2_causas)} causas en {t_v2:.1f}s")

        # Comparar
        comparar(v1_causas, v2_causas)

        print(f"\n  Tiempo v1 (Sonnet):  {t_v1:.1f}s")
        print(f"  Tiempo v2 (Haiku):   {t_v2:.1f}s")
        if t_v2 > 0:
            print(f"  Speedup:             {t_v1/t_v2:.1f}x")

        print(f"\n  Log comparativo: {log_path}")

    finally:
        sys.stdout = tee._stdout
        tee.close()


if __name__ == "__main__":
    main()
