# INCIDENTE WAF — Poder Judicial (2026-04-09/10)

## Resumen

El 09 de abril 2026, al implementar workers paralelos en M2, el WAF (Web Application Firewall F5 BIG-IP) del Poder Judicial comenzó a rechazar conexiones con el mensaje:
> "The requested URL was rejected. Please consult with your administrator."

## Cronología

### Día 1 (09 abril)
- Se implementaron 6 workers subprocess en `modulo2_ojv.py` (round-robin, stagger 3s entre lanzamientos)
- Primer run: 6 workers arrancaron pero todos recibieron "Request Rejected" de OJV
- Se asumió que era por múltiples conexiones simultáneas
- Se probó reducir a 5 workers → mismo error
- Se probó `channel="chrome"` (usar Chrome real en vez de Chromium bundled) → mismo error
- Se probó `--disable-blink-features=AutomationControlled` + `navigator.webdriver` override → mismo error
- Un abogado reportó que ni Firefox normal podía entrar a PDJ → **confirmó que era problema del servidor, no nuestro**
- Se probó CDP (Chrome DevTools Protocol) → no funcionó por problemas de perfil de Chrome
- Se probó `launch_persistent_context` con `channel="chrome"` → **CARGÓ PDJ exitosamente**
- Pero al navegar a `consultaunificadacausas.php` → "Request Rejected" de nuevo
- Se instaló `playwright-stealth` (pip install playwright-stealth) → **pasó el WAF completamente**

### Día 2 (10 abril)
- PDJ estaba estable nuevamente (el bloqueo del día anterior fue inestabilidad general del servidor)
- Test con 1 worker: 4/4 causas procesadas OK
- Test con 5 workers: 10/10 causas procesadas OK
- Pipeline completo (Regiones_docxToExcel.bat): 67 causas procesadas con 5 workers en ~7 min
- Filtrador (Detector_Excedentes.bat): 67 causas procesadas, 29 con "Error cuaderno" por carga de OJV
- Reaudit: 29 causas reprocesadas exitosamente (0 errores)

## Cambios técnicos aplicados

### En `ojv_remates.py`:
- `navegar_a_consulta()`: ahora va a `home/index.php` → click "Consulta causas" → selecciona competencia "Civil" → espera carga de opciones de Corte
- URL anterior: `indexN.php` (redirigía a login que el WAF bloqueaba)
- Agregado `playwright-stealth` (`Stealth().apply_stealth_sync(page)`)

### En `modulo2_ojv.py`:
- Workers subprocess con round-robin (no por Corte como el Filtrador)
- `launch_persistent_context` con `channel="chrome"` y perfil dedicado por worker (`.chrome-profile-wN`)
- Limpieza automática de perfiles al inicio de cada run (`shutil.rmtree`)
- Stagger de 3 segundos entre lanzamientos de workers
- Default: 5 workers, max: 10
- `playwright-stealth` aplicado a cada page

### En `filtrador_saldos.py`:
- `launch_persistent_context` con `channel="chrome"` (reemplazó `chromium.launch`)
- `playwright-stealth` aplicado
- Perfiles dedicados por worker

## Dependencia nueva
```
pip install playwright-stealth --break-system-packages
```

## Lecciones aprendidas

1. **PDJ tiene un WAF F5 BIG-IP** que detecta bots por fingerprint del browser (no solo por IP o rate limiting)
2. **`launch_persistent_context` + `channel="chrome"` + `playwright-stealth`** es la combinación que pasa el WAF de forma confiable
3. **Los perfiles persistentes deben limpiarse** antes de cada run porque pueden cachear estados de rechazo del WAF
4. **La inestabilidad de PDJ es real** — abogados confirman que históricamente el sitio tiene problemas frecuentes. No asumir que un rechazo es por nuestro código.
5. **`navegar_a_consulta` cambió de flujo**: antes iba directo a `indexN.php`; ahora pasa por `home/index.php` → click "Consulta causas" → selección de competencia Civil
6. **Reaudit es la herramienta para recuperar causas fallidas** — `--reaudit --workers 5` reprocesa PENDIENTE_ACTA sin perder trabajo previo
