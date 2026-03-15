# config_template.py — Copiar como config.py y rellenar con valores reales
# NO subir config.py a GitHub (está en .gitignore)

import os

# === RUTAS ===
BASE_DIR = r"D:\Remates"
DIARIOS_DIR = os.path.join(BASE_DIR, "Diarios")
DESCARGAS_DIR = os.path.join(BASE_DIR, "Descargas")
CAUSAS_XLSX = os.path.join(BASE_DIR, "causas_ojv.xlsx")

# === API KEYS ===
ANTHROPIC_API_KEY = "sk-ant-XXXXXXXXXX"  # Claude API (Sonnet 4.6 para M1)

# === CONSTANTES ===
CORTES_RM = ["C.A. de Santiago", "C.A. de San Miguel"]

DEMANDANTES_EXCLUIDOS = [
    "Banco Estado",
    "Banco del Estado",
]

# Causas con cuadernos restringidos/inaccesibles en OJV
CAUSAS_IGNORADAS = [
    # "C-1838-2024",  # ejemplo
]
