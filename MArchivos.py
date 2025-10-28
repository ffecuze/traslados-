import pandas as pd

# Archivos
ruta_maestro = r"C:\Users\pablo.larco\OneDrive - GR Chia S.A.S\Escritorio\FF\BI\CambiarArchivos\Maestro de empleados activos 26.06.2025 2.xlsx"
ruta_nomina = r"C:\Users\pablo.larco\OneDrive - GR Chia S.A.S\Escritorio\FF\BI\CambiarArchivos\Reporte nomina activa 28.05.2025 2.xls"

# Verificar hojas disponibles
xls_maestro = pd.ExcelFile(ruta_maestro)
print("Hojas en Maestro:", xls_maestro.sheet_names)

xls_nomina = pd.ExcelFile(ruta_nomina)
print("Hojas en Nómina:", xls_nomina.sheet_names)

# Leer la hoja correcta (ajusta según los nombres que imprima)
df_maestro = pd.read_excel(ruta_maestro, sheet_name="Hoja1")
df_nomina = pd.read_excel(ruta_nomina, sheet_name="Hoja1")

print(df_maestro.head())
print(df_nomina.head())
