# REPORT.IA RCV — Publicación en GitHub Pages

## Archivos principales
- `index.html`: interfaz principal.
- `styles.css`: diseño y estilos.
- `app.js`: lógica de carga, análisis, dashboard y exportaciones.
- `README.md`: información del proyecto.

## Cómo publicar
1. Crea un repositorio nuevo en GitHub.
2. Descomprime este ZIP.
3. Entra a la carpeta `REPORTIA_RCV`.
4. Sube **todos los archivos que están dentro de esa carpeta** a la raíz del repositorio.
5. En GitHub ve a **Settings > Pages**.
6. En **Build and deployment** selecciona **Deploy from a branch**.
7. Selecciona la rama `main` y la carpeta `/ (root)`.
8. Guarda los cambios.
9. GitHub mostrará la URL pública del portal.

## Importante
La aplicación procesa los Excel directamente en el navegador. No requiere servidor propio.
Las librerías externas se cargan mediante CDN, por lo que el equipo necesita acceso a internet para cargar dichas librerías.
