# REPORT.IA RCV 18.0 — Clean Rebuild

Esta versión reconstruye la estructura visual sin eliminar el motor funcional.

## Cambio principal
La aplicación ahora usa un contenedor `.v18-shell` con dos columnas reales:
- Sidebar
- Contenido principal

No usa `margin-left` para acomodar el contenido.

## Pantalla inicial
Solo muestra:
- Centro de Mando Ejecutivo
- KPIs
- Rendimiento
- Ranking
- Cumplimiento
- Desviaciones
- Hallazgos IA
- Accesos rápidos
- Copiloto

## Módulos anteriores
Se conservan, pero se muestran de uno en uno desde el menú lateral:
- Carga
- Análisis
- Comparativos
- Gerencias / Insights
- Reportes
- Configuración
- Centro de Decisiones
- Storytelling
- PDF

## Compatibilidad
Se conserva:
- Login existente
- Lógica de carga
- Generación FINAL
- Reportes especializados
- PDF
- Copiloto
- Apps Script / configuración existente
