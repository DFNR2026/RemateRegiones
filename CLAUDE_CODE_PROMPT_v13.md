## ENTORNO DE EJECUCIÓN — PRIORIDAD LOCAL (pactado 2026-05)

La ejecución de código se delega prioritariamente a Qwen3-Coder:30b LOCAL en
Roo Code/Ollama (MoE 30B/3.3B activos; hardware: RTX 5070 Ti 16GB, Ryzen 5 9600X,
32GB DDR5). Claude Code queda en segundo plano (solo si el local se queda corto).

Reparto de roles:
- Claude (chat web): arquitecto de diseño y diagnóstico semántico. Entrega
  instrucciones acotadas con puntaje Qwen 1-10 por tarea. No escribe scripts largos.
- Qwen local: ejecutor. Escribe módulos, openpyxl, dry-runs. Lee .clinerules.
- Diego: árbitro. Corre y valida; el local no decide correctitud.

Las reglas de ejecución para el modelo local viven en D:\Remates\.clinerules
(incluye la estructura modular objetivo del filtrador). Mantener ese archivo
sincronizado con cualquier cambio de arquitectura.

Extrapolación al proyecto RM: PAUSADA por completo (otro repositorio, se retoma al final).