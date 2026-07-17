# REPORT.IA RCV · Intelligence Edition 2.1 Secure Executive Stable

Esta versión combina:
- Todas las funciones de Intelligence Edition 2.1.
- Inicio de sesión.
- Lista desplegable de usuarios.
- Compatibilidad `fetch` + JSONP.
- Corrección de PDF con modo automático y modo Guardar como PDF.

## Apps Script
Usa el archivo `APPS_SCRIPT_LOGIN_EJEMPLO.gs`.

Después de actualizar el script:
1. Implementar > Administrar implementaciones.
2. Edita la implementación actual.
3. Selecciona Nueva versión.
4. Ejecutar como: Tú.
5. Acceso: Cualquier persona.
6. Implementar.

## Pruebas
Usuarios:
`?accion=usuarios`

JSONP:
`?accion=usuarios&callback=prueba`

Debe verse:
`prueba({"ok":true,"usuarios":[...]});`

## GitHub
Reemplaza todos los archivos de la versión anterior por los de esta carpeta.
