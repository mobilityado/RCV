# REPORT.IA RCV · Login estilo AVA

Esta versión simplifica completamente el acceso.

## Flujo
1. Intenta cargar usuarios desde Apps Script:
   `https://script.google.com/macros/s/AKfycbxUhENeMAGaVJx2Gs4yR_qncJJxHyq8NlFFSfa9qu7XBDgcDu4L9HasfrzZSQBOKgwp/exec?accion=usuarios`
2. Si el navegador no permite leer esa respuesta, intenta cargar los nombres directamente desde Google Sheets público.
3. Si ambos fallan, usa una lista de nombres de respaldo.
4. La contraseña SIEMPRE se valida contra Apps Script.

## Apps Script
Usa `APPS_SCRIPT_AVA_STYLE.gs`.

Luego:
- Implementar > Administrar implementaciones.
- Nueva versión.
- Ejecutar como: Tú.
- Acceso: Cualquier persona.

## Nota
La lista de respaldo solo contiene nombres y roles. Nunca contiene contraseñas.
