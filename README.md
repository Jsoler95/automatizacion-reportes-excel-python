# Automatización de Reportes Excel con Python

Este proyecto automatiza la creación de reportes a partir de varios archivos Excel. El programa lee múltiples archivos de ventas, los une en una sola tabla, limpia los datos, calcula el total de ventas y genera un reporte final en Excel con varias hojas.

## Problema que resuelve

Muchas empresas manejan reportes mensuales separados en diferentes archivos Excel. Unir esos archivos manualmente puede tomar tiempo y provocar errores. Este script automatiza ese proceso para ahorrar tiempo y mejorar la organización de los datos.

## Funcionalidades

- Lee automáticamente todos los archivos `.xlsx` dentro de la carpeta `ventas/`.
- Une todos los archivos en una sola tabla.
- Limpia espacios en columnas de texto.
- Estandariza el formato de la ciudad.
- Convierte columnas numéricas como `cantidad` y `precio`.
- Elimina filas con datos incompletos o inválidos.
- Calcula una nueva columna llamada `total`.
- Agrega una columna `archivo_origen` para identificar de qué archivo proviene cada fila.
- Genera un archivo Excel final con varias hojas:
  - `Datos completos`
  - `Resumen por ciudad`
  - `Resumen por producto`

## Tecnologías utilizadas

- Python
- pandas
- openpyxl
- pathlib

## Estructura del proyecto

```text
automatizacion-reportes-excel-python/
│
├── reporte_ventas.py
├── ejecutar_reporte.bat
├── README.md
├── requirements.txt
├── ventas/
│   ├── ventas_enero.xlsx
│   ├── ventas_febrero.xlsx
│   └── ventas_marzo.xlsx
└── output/
    └── reporte_final.xlsx
```

## Cómo ejecutar el proyecto

1. Clona este repositorio o descarga los archivos.

2. Instala las dependencias:

```bash
py -m pip install -r requirements.txt
```

3. Coloca los archivos Excel dentro de la carpeta `ventas/`.

4. Ejecuta el proyecto con una de estas opciones:

**Opción 1: desde la terminal**

```bash
py reporte_ventas.py
```

**Opción 2: en Windows usando el archivo .bat**

Haz doble clic en:

```text
ejecutar_reporte.bat
```

Este archivo ejecuta el script automáticamente y abre el reporte final cuando termina.

5. El reporte final se generará automáticamente en:

```text
output/reporte_final.xlsx
```