# Visor Excel Interactivo

Aplicación web (HTML + JS) para arrastrar un archivo Excel y extraer rápidamente los datos más relevantes de la primera hoja:

- Resumen general (archivo, hoja, total de filas y columnas).
- Análisis por columna:
  - Tipo detectado (numérica o texto/mixta).
  - Completitud (% de celdas no vacías).
  - Dato destacado:
    - Numérica: mínimo, máximo y promedio.
    - Texto/mixta: valor más frecuente.
- Vista previa con las primeras 10 filas.

## Uso

1. Abre `index.html` en tu navegador.
2. Arrastra un `.xlsx` o `.xls` al área central (o haz clic para elegir archivo).
3. Revisa automáticamente el resumen y los indicadores relevantes.

## Nota

Se usa SheetJS desde CDN para leer archivos Excel en el navegador.
