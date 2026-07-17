# REPORT.IA RCV 18.1 — Fresh Frontend

Esta versión separa por completo la interfaz visible del motor heredado.

- El motor original permanece oculto en `#legacyEngineRoot`.
- El frontend visible usa clases `fresh-*` independientes.
- Los estilos antiguos ya no controlan la geometría del dashboard visible.
- La carga de archivos y generación FINAL siguen usando el motor existente.
- El dashboard nuevo lee y refleja la información procesada.
- Los reportes especializados se abren desde la nueva interfaz.
- El login existente se conserva.

Archivos nuevos:
- v18-fresh.css
- v18-fresh.js
