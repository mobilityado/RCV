# REPORT.IA RCV — Intelligence Executive 8.0

Versión mejorada sobre el proyecto original RCV, manteniendo su procesamiento local de Excel y sus módulos existentes.

## Mejoras principales

- Rediseño ejecutivo más compacto y consistente.
- Sidebar oscuro premium con estado del sistema.
- Encabezado superior tipo aplicación ejecutiva.
- Pantalla breve de arranque.
- Login reforzado con mensajes de error claros cuando Apps Script devuelve HTML, expira o no responde.
- Compatibilidad con respuestas de autenticación que usen `usuarios` o `users`, y roles `tipo`, `rol` o `role`.
- Tarjetas KPI y gráficas con proporciones más equilibradas.
- Mejor comportamiento responsive.
- Se mantiene intacta la lógica de carga de archivos, generación de reportes, ZIP, análisis, gerencias, costos, gastos, XPV, tendencias y comparador.

## Instalación en GitHub Pages

1. Sube todos los archivos de esta carpeta a la raíz del repositorio.
2. Verifica en `config.js` que `authEndpoint` tenga la URL `/exec` vigente de tu Apps Script.
3. Activa GitHub Pages desde la rama `main`, carpeta `/root`.
4. Abre la página y usa Ctrl+F5 una vez después de publicar.

## Autenticación

El proyecto sigue usando el Apps Script configurado en `config.js`. No se modificó la estructura esperada de tu servicio; solo se mejoró la tolerancia a errores y formatos de respuesta.
