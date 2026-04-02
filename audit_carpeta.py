"""
audit_carpeta.py -- Auditoria recursiva de D:\Remates\

Genera un log con todos los archivos, tamanos, fechas, y alertas de sospechosos.

Uso: python audit_carpeta.py
"""

import os
import datetime

BASE_DIR = r"D:\Remates"
LOGS_DIR = os.path.join(BASE_DIR, "logs")
os.makedirs(LOGS_DIR, exist_ok=True)

LOG_FILE = os.path.join(
    LOGS_DIR,
    f"audit_carpeta_{datetime.datetime.now().strftime('%Y%m%d')}.log",
)

# Scripts .py conocidos (sin ruta, solo nombre)
KNOWN_PY = {
    "main.py", "config.py", "ojv_remates.py", "filtrador_saldos.py",
    "limpiar_cache.py", "audit_carpeta.py",
    "modulo1_parser.py", "modulo2_ojv.py", "modulo3_extractor.py",
    "modulo4_tasacion.py", "modulo5_reporte.py",
    # v2
    "modulo1_v2.py",
}

# Prefijos conocidos
KNOWN_PY_PREFIXES = ("fix_", "diagnostico_")

EXCLUDE_DIRS = {".git", "__pycache__", "node_modules", ".venv", "venv"}


def _format_size(size_bytes):
    if size_bytes >= 1024 * 1024:
        return f"{size_bytes / (1024 * 1024):.1f} MB"
    return f"{size_bytes / 1024:.1f} KB"


def _is_known_py(filename):
    if filename in KNOWN_PY:
        return True
    for prefix in KNOWN_PY_PREFIXES:
        if filename.startswith(prefix):
            return True
    return False


def main():
    lines = []
    ext_count = {}
    ext_size = {}
    folder_size = {}
    sospechosos = []
    pdfs_pequenos = []
    temp_viejos = []
    logs_grandes = []

    now = datetime.datetime.now()
    hace_30_dias = now - datetime.timedelta(days=30)

    for dirpath, dirnames, filenames in os.walk(BASE_DIR):
        # Excluir directorios
        dirnames[:] = [d for d in dirnames if d not in EXCLUDE_DIRS]

        if not filenames:
            continue

        lines.append(f"=== DIRECTORIO: {dirpath} ===")

        rel_dir = dirpath  # ruta completa para folder_size
        if rel_dir not in folder_size:
            folder_size[rel_dir] = 0

        for fname in sorted(filenames):
            fpath = os.path.join(dirpath, fname)
            try:
                stat = os.stat(fpath)
                fsize = stat.st_size
                fdate = datetime.datetime.fromtimestamp(stat.st_mtime)
            except OSError:
                continue

            size_str = _format_size(fsize)
            date_str = fdate.strftime("%d/%m/%Y")

            lines.append(f"  {fname:<50s} {size_str:>10s}   {date_str}")

            # Contadores
            _, ext = os.path.splitext(fname)
            ext = ext.lower() if ext else "(sin ext)"
            ext_count[ext] = ext_count.get(ext, 0) + 1
            ext_size[ext] = ext_size.get(ext, 0) + fsize
            folder_size[rel_dir] = folder_size.get(rel_dir, 0) + fsize

            # Alertas
            if ext == ".py" and not _is_known_py(fname):
                sospechosos.append(f"{fpath}  <-- script desconocido")

            if ext == ".pdf" and "liquidaciones" in dirpath.lower() and fsize < 100 * 1024:
                pdfs_pequenos.append(f"{fpath} ({_format_size(fsize)})  <-- posible resolucion, no liquidacion")

            if "temp_workers" in dirpath and fdate < hace_30_dias:
                temp_viejos.append(f"{fpath} ({date_str})  <-- temporal antiguo (>30 dias)")

            if ext == ".log" and fsize > 10 * 1024 * 1024:
                logs_grandes.append(f"{fpath} ({_format_size(fsize)})  <-- log grande (>10MB)")

        lines.append("")

    # Resumen
    lines.append("=" * 60)
    lines.append("  RESUMEN")
    lines.append("=" * 60)

    lines.append("")
    lines.append("  Archivos por extension:")
    for ext in sorted(ext_count, key=lambda e: ext_count[e], reverse=True):
        lines.append(f"    {ext:<12s} {ext_count[ext]:>4d} archivos  ({_format_size(ext_size[ext])})")

    lines.append("")
    lines.append("  Espacio por carpeta:")
    for folder in sorted(folder_size, key=lambda f: folder_size[f], reverse=True):
        if folder_size[folder] > 0:
            lines.append(f"    {folder:<60s} {_format_size(folder_size[folder])}")

    # Alertas
    all_alerts = []
    all_alerts.extend(sospechosos)
    all_alerts.extend(pdfs_pequenos)
    all_alerts.extend(temp_viejos)
    all_alerts.extend(logs_grandes)

    if all_alerts:
        lines.append("")
        lines.append("  SOSPECHOSOS:")
        for alert in all_alerts:
            lines.append(f"    {alert}")

    lines.append("")
    lines.append("=" * 60)

    total_files = sum(ext_count.values())
    total_size = sum(folder_size.values())
    lines.append(f"  Total: {total_files} archivos, {_format_size(total_size)}")
    lines.append(f"  Log guardado en: {LOG_FILE}")
    lines.append("=" * 60)

    output = "\n".join(lines)

    # Consola
    print(output)

    # Archivo
    with open(LOG_FILE, "w", encoding="utf-8") as f:
        f.write(output)


if __name__ == "__main__":
    main()
