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

## Premium Dashboard 3.0
Rediseño visual completo inspirado en dashboards ejecutivos BI:
- Sidebar fijo oscuro con navegación.
- Carga de archivos compacta integrada en el panel lateral.
- Filtros superiores.
- KPIs en una línea ejecutiva.
- Gráficas y tablas en una cuadrícula consistente.
- Centro de Reportes Ejecutivos en tarjetas.
- PDFs, Excel Ejecutivo y paquete ZIP.
- Diseño responsive y modo presentación.


## Versión 3.1 — Diseño final de publicación
- Sidebar inspirado en el mockup ejecutivo aprobado.
- Carga de archivos compacta y discreta.
- Dashboard visible desde el primer momento.
- Navegación más sobria y profesional.
- Preparado para GitHub Pages.

## Mockup Style 3.2
- Reestructuración visual para coincidir con el mockup aprobado.
- Sidebar más ancho y respirado.
- Carga de archivos integrada de forma compacta.
- Dashboard visible desde arriba.
- KPIs y alertas alineados.
- Tendencias, mapa por gerencia y reportes ejecutivos en cuadrícula fija.
- Preparado para GitHub Pages.


## v3.3 — Carga en área de trabajo
- Se retiró la carga de archivos del sidebar.
- La carga ahora aparece como una tarjeta compacta en el área blanca superior.
- El sidebar queda dedicado únicamente a navegación y usuario.

## REPORT.IA RCV · Pro Edition 4.0
Nuevas funciones:
- Explorador drill-down: gerencia → agrupador → cuenta.
- Historial local de periodos usando LocalStorage.
- Comparador ejecutivo enriquecido.
- Módulo "Qué cambió" integrado con el análisis.
- Informe ejecutivo automático con narrativa y PDF.
- Barra de accesos rápidos para análisis avanzado.

## REPORT.IA RCV · Pro Edition 5.0
Nuevas funciones:
- Centro de Prioridades con focos críticos y oportunidades.
- Copiloto REPORT.IA con preguntas en lenguaje natural sobre los datos cargados.
- Cierre Ejecutivo automático del periodo.
- Modo Comité para presentación limpia de Resumen, Prioridades, Gerencias, Tendencias y Conclusión.


## Enterprise Edition 6.0
- Login conectado a Google Apps Script.
- Sesión activa y cierre de sesión.
- Saludo personalizado y rol de usuario.
- Home Ejecutiva simplificada.
- Navegación por módulos para evitar saturación visual.
- Archivo de ejemplo de Apps Script para autenticación.

## Enterprise Edition 6.1
- Selector desplegable de usuarios cargado desde Google Sheets vía Apps Script.
- Avatar dinámico con iniciales del usuario seleccionado.
- Contraseña permanece separada y nunca se lista desde la hoja.
- Apps Script corregido para evitar el error `e.parameter` al ejecutar manualmente `doGet`.

## Enterprise Edition 6.2
- Corregida la carga de usuarios desde GitHub Pages.
- Intenta `fetch` y usa JSONP automáticamente como respaldo.
- Apps Script actualizado con soporte `callback`.
- Botón para recargar usuarios en la pantalla de acceso.

## Signature Edition 7.0
- Rediseño visual premium.
- Hero ejecutivo y Pulso RCV circular.
- Microinteracciones y sparklines.
- Sidebar colapsable.
- Modo oscuro.
- Command Palette.
