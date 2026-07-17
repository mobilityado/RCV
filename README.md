# REPORT.IA RCV

Aplicación web estática para automatizar la generación de reportes RCV desde archivos JD.

## Funciones

- Acepta un solo Excel con las pestañas:
  - `CATALOGO DE GERENTES`
  - `JD COSTO RCV`
  - `JD XPV RCV`
  - `JD GASTOS RCV`
- También acepta los archivos por separado.
- Detecta y valida automáticamente las fuentes.
- Permite seleccionar el mes de corte.
- Genera:
  - `COSTO RCV 2026.xlsx`
  - `Gastos RCV 2026.xlsx`
  - `Productividad XPV [mes] 26 RCV.xlsx`
  - `FINAL.zip`
- Todo se procesa localmente en el navegador.

## Publicar en GitHub Pages

1. Crea un repositorio nuevo en GitHub.
2. Sube `index.html`, `styles.css` y `app.js` a la raíz.
3. En **Settings → Pages**, selecciona **Deploy from a branch**.
4. Selecciona la rama `main` y la carpeta `/ (root)`.
5. Guarda y abre la URL publicada por GitHub Pages.

## Dependencias

La aplicación carga SheetJS y JSZip desde CDN. Por ello, al abrirla se requiere conexión a internet para cargar esas dos librerías. Los archivos Excel del usuario no se envían a esas librerías ni a un servidor: el procesamiento ocurre en el navegador.

## Lógica actual de generación

Para Costos y Gastos se toman los campos de JD y se acumulan hasta el mes seleccionado:

- 2025 = `REAL GESTION` del año 2025.
- PTTO. 2026 = `PRESUPUESTO GESTION` del año 2026.
- 2026 = `REAL GESTION` del año 2026.
- Se calculan diferencias 25 vs 26 y PTTO vs 26.
- Se aplica el catálogo de gerentes cuando está disponible.

En Costos se generan las hojas `COSTOS OPERATIVOS`, `COSTO MANTTO`, `premisas 26` y `SMO` con reglas de clasificación incorporadas.

## Nota de compatibilidad con el FINAL histórico

Esta versión reconstruye la lógica de cálculo directamente en JavaScript y no depende del complemento XLCubed. Debido a que los archivos históricos contienen pivotes/conexiones y estructuras propias de Excel/XLCubed, pueden existir diferencias de presentación o de reglas específicas frente al archivo histórico. La app está preparada para ajustar esas reglas en `app.js` sin cambiar la interfaz.
