# REPORT.IA RCV · PDF Fix

## Qué se corrigió
La versión anterior podía mostrar:

`No se pudieron cargar las librerías de PDF`

porque `html2canvas` y `jsPDF` no estaban disponibles.

La versión corregida usa dos modos:

### Modo 1 — PDF automático
Si el navegador puede cargar `html2canvas` y `jsPDF`, REPORT.IA descarga el PDF directamente.

### Modo 2 — PDF compatible
Si la red corporativa bloquea esas librerías, REPORT.IA abre una vista de impresión limpia.

En Chrome/Edge:
1. Pulsa `PDF de la vista actual`.
2. Se abrirá el diálogo de impresión.
3. En Destino selecciona `Guardar como PDF`.
4. Pulsa Guardar.

Este modo no depende de las librerías externas de PDF.

## También se corrigió
- Se eliminó el `alert()` invasivo.
- Ahora se muestran notificaciones integradas dentro del portal.
- El Paquete Ejecutivo ya no falla si `html2canvas` no está disponible.
