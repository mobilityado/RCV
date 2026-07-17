# REPORT.IA RCV · Enterprise Edition 6.0

## Inicio de sesión
El portal está configurado para consultar:

https://script.google.com/macros/s/AKfycbxUhENeMAGaVJx2Gs4yR_qncJJxHyq8NlFFSfa9qu7XBDgcDu4L9HasfrzZSQBOKgwp/exec

La interfaz envía:
- `accion=login`
- `usuario=<usuario>`
- `contrasena=<contraseña>`

Respuesta esperada:
```json
{"ok": true, "usuario": "NOMBRE", "tipo": "ADMINISTRADOR"}
```

Si tu Apps Script actual no implementa esa respuesta, copia el contenido de `APPS_SCRIPT_LOGIN_EJEMPLO.gs` en tu proyecto de Apps Script y vuelve a implementar la aplicación web.

## Nueva Home Ejecutiva
Después del login, el usuario entra primero a una pantalla limpia con:
- Pulso RCV.
- Prioridades.
- Oportunidad del periodo.
- Accesos a Análisis, Gerencias, Explorador y Reportes.

## Navegación
Cada opción del menú abre su módulo de forma independiente. Ya no se muestran todas las gráficas al mismo tiempo.

## Sesión
La sesión se mantiene mientras la pestaña del navegador permanezca abierta mediante `sessionStorage`.

## Publicación
Sube todos los archivos de la carpeta `REPORTIA_RCV` a la raíz de tu repositorio y activa GitHub Pages.


## Enterprise Edition 6.1 — Selector de usuarios
La pantalla de login ahora consulta automáticamente:

`?accion=usuarios`

Respuesta esperada:

```json
{
  "ok": true,
  "usuarios": [
    {
      "usuario": "JOSE ANGEL LOPEZ HERNANDEZ",
      "tipo": "ADMINISTRADOR"
    }
  ]
}
```

Los nombres aparecen en una lista desplegable. La contraseña no se descarga ni se muestra; solo se valida cuando el usuario intenta iniciar sesión.

Después de modificar Apps Script:
1. Guarda los cambios.
2. Ve a Implementar > Administrar implementaciones.
3. Edita la implementación.
4. Selecciona Nueva versión.
5. Implementa nuevamente.
