# REPORT.IA RCV — Intelligence Decision 10.1 PDF FIX

Corrección del exportador PDF:

- Captura el contenido visible después de cargar fuentes e imágenes.
- Divide físicamente el canvas en una imagen independiente por página.
- Evita la técnica anterior de mover una imagen gigante con coordenadas negativas.
- Oculta botones, modales y herramientas flotantes durante la captura.
- Agrega encabezado, fecha y número de página.
- Mantiene alternativa Imprimir → Guardar como PDF si la captura falla.

No se modifica la lógica original de datos, autenticación, alertas, Decision Center ni drill-down.
