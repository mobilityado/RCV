# REPORT.IA RCV — Intelligence Decision 10.4 · Generator Studio

## Propósito
REPORT.IA RCV automatiza la transformación de los reportes JD fuente en el paquete de reportes FINAL.

### Entradas detectadas
- CATALOGO DE GERENTES
- JD COSTO RCV
- JD XPV RCV
- JD GASTOS RCV

Pueden cargarse:
1. En un único Excel, cada fuente en su propia pestaña.
2. Como archivos Excel independientes.

### Salidas generadas
- COSTO RCV 2026.xlsx
  - COSTOS OPERATIVOS
  - COSTO MANTTO
  - premisas 26
  - SMO
- Gastos RCV 2026.xlsx
  - GASTOS RCV
  - RESUMEN GERENCIA
- Productividad XPV [Mes] 26 RCV.xlsx
  - PRODUCTIVIDAD XPV

El procesamiento se realiza localmente en el navegador mediante JavaScript y SheetJS.

## Cambio visual 10.4
La pantalla inicial ahora tiene:
- Descripción del propósito del proyecto a la izquierda.
- Flujo de trabajo Entrada → Procesamiento → FINAL.
- Resumen de los tres reportes generados.
- Carga de archivos compacta a la derecha.
- Mes de corte y botón de generación dentro del mismo bloque.
