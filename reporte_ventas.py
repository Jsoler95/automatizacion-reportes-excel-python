# Importo las librerias que voy a estar utilizando
import pandas as pd
from pathlib import Path

carpeta_proyecto = Path(__file__).parent

carpeta_ventas = carpeta_proyecto / "ventas"
carpeta_output = carpeta_proyecto / "output"

# Creo la carpeta output 
carpeta_output.mkdir(exist_ok=True)

# Creo una lista vacia para guardar las tablas de los archivos Excel
tablas = []

# Lee todos los archivos dentro de la carpeta ventas
for archivo in carpeta_ventas.glob("*xlsx"):
    df = pd.read_excel(archivo)

    df["archivo_origen"] = archivo.name

    #Guardo las tablas dentro de la lista
    tablas.append(df)

# Uno las tablas en una tabla nueva
df_final = pd.concat(tablas, ignore_index= True)

# Limpio las columnas de texto
df_final["producto"] = df_final["producto"].str.strip().str.lower()
df_final["precio"] = pd.to_numeric(df_final["precio"], errors="coerce")

# Elimino filas con datos importantes vacios o invalidos
df_final = df_final.dropna(subset = ["producto","cantidad","precio","ciudad"])

# Creo una nueva columna con el total
df_final["total"] = df_final["cantidad"] * df_final["precio"]

# Creo un resumen por ciudad y producto
resumen_ciudad = df_final.groupby("ciudad")["total"].sum().reset_index()
resumen_producto = df_final.groupby("producto")["total"].sum().reset_index()

# Ruta del reporte
archivo_salida = carpeta_output / "reporte_final.xlsx"

# Creo archivo Excel con varias hojas
with pd.ExcelWriter(archivo_salida, engine="openpyxl") as writer:
    df_final.to_excel(writer,sheet_name = "Datos completos", index=False)
    resumen_ciudad.to_excel(writer,sheet_name = "Resumen por ciudad", index=False)
    resumen_producto.to_excel(writer,sheet_name = "Resumen por producto", index=False)

print(f"Reporte creado correctamente: {archivo_salida}")