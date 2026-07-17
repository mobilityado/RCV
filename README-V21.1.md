# REPORT.IA RCV 21.1 — Universal Loader

Modos de entrada soportados:

## Modo A — Concentrado reportes JD
Un solo Excel (o varios archivos) que en conjunto contengan:
- CATALOGO DE GERENTES
- JD COSTO RCV
- JD GASTOS RCV
- JD XPV RCV

## Modo B — Archivos RP
Puede seleccionar los archivos sueltos o un ZIP:
- COSTO RCV
- Gastos RCV
- Productividad XPV

REPORT.IA detecta automáticamente el modo de entrada y procesa ambos caminos hacia el mismo dashboard y Centro de Reportes.

Nota:
En modo RP se leen directamente las tablas ya estructuradas de Costos, Gastos y Productividad. En modo JD se conserva el motor de transformación validado de la versión 19.
