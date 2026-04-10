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
  - Limpieza de filtros con un clic.
- Gráfico rápido (barras): top 8 valores más frecuentes de la columna seleccionada.
- Vista previa con las primeras 10 filas filtradas.

## Uso

1. Abre `index.html` en tu navegador.
2. Arrastra un `.xlsx` o `.xls` al área central (o haz clic para elegir archivo).
3. Aplica filtros por columna/valor y revisa cómo cambian análisis, gráfico y vista previa.

## Nota

Se usa SheetJS desde CDN para leer archivos Excel en el navegador.
