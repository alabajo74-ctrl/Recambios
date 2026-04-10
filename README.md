# Visor Excel Interactivo

Aplicación web (HTML + JS) para arrastrar un archivo Excel y extraer rápidamente los datos más relevantes de la primera hoja.

## Funcionalidades

- Carga de archivo por drag & drop o selector (`.xlsx` y `.xls`).
- Resumen general (archivo, hoja, total de filas y columnas).
- Análisis por columna:
  - Tipo detectado (numérica o texto/mixta).
  - Completitud (% de celdas no vacías).
  - Dato destacado:
    - Numérica: mínimo, máximo y promedio.
    - Texto/mixta: valor más frecuente.
- Filtros interactivos:
  - Filtro por columna específica o por todas.
  - Búsqueda por texto/valor.
  - Filtro dedicado por almacén (si existe columna con "almacén").
  - Filtro dedicado facturable/no facturable (si existe columna con "facturable").
- Gráficos automáticos de negocio:
  - Barras por almacén (top 8).
  - Facturable vs no facturable (gráfico circular).
- Gráficos de insights relevantes:
  - Completitud por columna (top 8).
  - Distribución de tipos de columna (numérica, texto/mixta, vacía).
- Gráfico rápido adicional (barras): top 8 valores más frecuentes de cualquier columna seleccionada.
- Vista previa con las primeras 10 filas filtradas.

## Uso

1. Abre `index.html` en tu navegador.
2. Arrastra un `.xlsx` o `.xls` al área central (o haz clic para elegir archivo).
3. Aplica filtros por columna/valor, por almacén y por facturable para actualizar análisis y gráficos.

## Nota

Se usa SheetJS desde CDN para leer archivos Excel en el navegador.
