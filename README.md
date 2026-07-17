# REPORT.IA RCV · Executive Suite

Portal web 100% del lado del navegador para cargar reportes JD, analizarlos y generar reportes RCV.

## Módulos

- Carga de Excel consolidado por pestañas o archivos separados.
- Detección automática de JD COSTO RCV, JD XPV RCV, JD GASTOS RCV y catálogo de gerentes.
- Resumen Ejecutivo con KPIs y desviaciones.
- Análisis por Gerencia con ficha ejecutiva, semáforo, tendencia, drivers y XPV.
- Módulo de Costos.
- Módulo de Gastos.
- Módulo de Productividad XPV.
- Tendencias acumuladas 2025 / Presupuesto 2026 / Real 2026.
- Resumen Inteligente y mapa ejecutivo de gerencias.
- Descarga individual de reportes y paquete FINAL.zip.

## Publicar en GitHub Pages

1. Crea un repositorio nuevo en GitHub.
2. Sube `index.html`, `styles.css`, `app.js`, `README.md` y `.gitignore` a la raíz.
3. En Settings > Pages selecciona Deploy from a branch.
4. Selecciona la rama `main` y carpeta `/root`.
5. Guarda y abre la URL publicada.

## Privacidad

El procesamiento se realiza localmente en el navegador. Los archivos cargados no se envían a un servidor por esta aplicación.

## Nota

La lógica de generación reconstruye los cálculos en JavaScript y no depende de complementos externos como XLCubed. Para igualar al 100% reportes históricos específicos, pueden requerirse ajustes adicionales a reglas de negocio o formatos particulares de las plantillas originales.


## Intelligence Edition
Esta edición añade Pulso RCV 0-100, centro de alertas, bloque “Qué cambió”, comparador de gerencias y modo presentación. Las funciones trabajan sobre los datos procesados localmente y no requieren backend.


## Intelligence Edition 2.0
- Exportación de la vista actual a PDF.
- Exportación PDF por gerencia seleccionada.
- Excel Ejecutivo de KPIs visibles.
- Paquete Ejecutivo ZIP.
- Paleta corporativa refinada: morado profundo, blanco/gris y colores semánticos para alertas.

Nota: la reproducción 1:1 de plantillas Excel con componentes propietarios (por ejemplo XLCubed) requiere validar cada plantilla original y sus conexiones; esta versión conserva el generador operativo existente y añade exportaciones ejecutivas del portal.

## Intelligence Edition 2.1 — Reorganización ejecutiva
- Jerarquía visual renovada.
- Barra superior de identidad ejecutiva.
- Filtros y formularios agrupados.
- KPIs alineados en cuadrícula consistente.
- Gráficas organizadas con espaciado uniforme.
- Botones de acción y descarga agrupados.
- Centro de Reportes Ejecutivos claramente separado.
- Mejor comportamiento en tablet y celular.
- Modo presentación más limpio.


## Intelligence Edition 2.1 Secure Executive
- Base funcional restaurada desde Intelligence Edition 2.1.
- Login integrado con Apps Script.
- Vista ejecutiva premium sin eliminar ningún módulo original.
