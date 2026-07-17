# REPORT.IA RCV · Enterprise Edition 6.2

## Por qué el JSON abre en el navegador pero el desplegable puede no cargar

La URL de Apps Script funciona correctamente, pero una página alojada en GitHub Pages puede encontrarse con restricciones de CORS o con la redirección de `script.google.com` a `script.googleusercontent.com`.

La versión 6.2 intenta primero una carga normal mediante `fetch`. Si el navegador no permite esa respuesta, utiliza automáticamente JSONP.

## Paso obligatorio

Reemplaza tu Apps Script por el contenido de:

`APPS_SCRIPT_LOGIN_EJEMPLO.gs`

Después:

1. Guarda el proyecto.
2. Implementar > Administrar implementaciones.
3. Editar la implementación actual.
4. Selecciona **Nueva versión**.
5. Ejecutar como: **Tú**.
6. Acceso: **Cualquier persona**.
7. Implementa.

La URL `/exec` puede conservarse si editas la misma implementación.

## Pruebas

JSON normal:

`?accion=usuarios`

JSONP:

`?accion=usuarios&callback=prueba`

La segunda debe mostrar algo parecido a:

`prueba({"ok":true,"usuarios":[...]});`

Una vez que ambas funcionan, GitHub Pages podrá llenar el desplegable.

## Botón Recargar usuarios

La versión 6.2 incluye un botón `Recargar usuarios` en el login para repetir la conexión sin tener que recargar toda la página.
