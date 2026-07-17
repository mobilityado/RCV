# REPORT.IA RCV 19.0 — Validated Engine

Esta versión fue reconstruida desde cero usando como referencia los archivos reales entregados:

- `reportes JD(1).xlsx`
- `FINAL(1).zip`

## Motor validado

### COSTO RCV 2026
El motor acumula desde enero hasta el periodo seleccionado:
- 2025 = columna H `REAL GESTION`
- PTTO. 2026 = columna K `PRESUPUESTO GESTION`
- 2026 = columna J `REAL GESTION`

Agrupa por:
- Jerarquía Cuenta Contable
- CIA / Marca
- Cuenta Contable
- Gerencia mapeada desde `CATALOGO DE GERENTES`
- Centro de Gestión

Genera:
- COSTOS OPERATIVOS
- COSTO MANTTO
- premisas 26
- SMO

### Gastos RCV
Usa la misma lógica acumulada y genera:
- GASTOS RCV
- RESUMEN GERENCIA

### Productividad XPV
Agrupa por:
- Gerencia
- Cuenta contable
- Centro / Sublibro
- CIA / Marca

Genera:
- PRODUCTIVIDAD XPV

## Importante
Los KPIs del dashboard se calculan directamente desde el mismo modelo que genera los Excel FINAL. No se usan valores de demostración ni scraping del HTML.

## Validación
La pantalla `Cargar reportes` permite seleccionar opcionalmente un `FINAL.zip` de referencia. REPORT.IA compara:
- cantidad de filas
- total 2026 / real
de los reportes principales.

## Uso
1. Sube todo el contenido a GitHub Pages.
2. Abre REPORT.IA.
3. Ve a `Cargar reportes`.
4. Selecciona un Excel consolidado o varios Excel.
5. Procesa la información.
6. Descarga los tres archivos FINAL o el ZIP completo.

Las librerías SheetJS y JSZip se cargan desde jsDelivr.
