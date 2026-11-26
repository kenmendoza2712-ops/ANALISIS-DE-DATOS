from datetime import datetime
from pathlib import Path
import sqlite3
import pandas as pd
import matplotlib.pyplot as plt # Para el grafico 
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference

# Configuración de rutas: cada ejecución crea su propia carpeta nueva junto al script
BASE = Path(__file__).resolve().parent
marca = datetime.now().strftime("%Y%m%d_%H%M%S")   # ej: 20261126_143015
CARPETA = BASE / f"Salida_{marca}"
CARPETA.mkdir(exist_ok=False)                       # falla si ya existe, nunca pisa una corrida
DB = BASE / "tienda_retro.db"

ruta_png = CARPETA / "ventas_por_categoria.png"
ruta_reporte = CARPETA / "reporte_ventas_categoria.csv"
ruta_limpias = CARPETA / "ventas_limpias.csv"
ruta_excel = CARPETA / "reporte_ventas.xlsx"


def dar_formato(ws, columnas_moneda=()):
    """Aplica diseño a una hoja de Excel: encabezado, bordes, filas alternadas, anchos y formato de moneda."""
    azul = PatternFill("solid", fgColor="1F4E78")
    gris = PatternFill("solid", fgColor="F2F2F2")
    lado = Side(style="thin", color="BFBFBF")
    borde = Border(left=lado, right=lado, top=lado, bottom=lado)

    encabezados = [c.value for c in ws[1]]
    idx_moneda = {i + 1 for i, nombre in enumerate(encabezados) if nombre in columnas_moneda}

    # Encabezado
    for celda in ws[1]:
        celda.fill = azul
        celda.font = Font(bold=True, color="FFFFFF")
        celda.alignment = Alignment(horizontal="center", vertical="center")
        celda.border = borde

    # Cuerpo: bordes, filas alternadas y formato de moneda
    for fila in ws.iter_rows(min_row=2):
        for celda in fila:
            celda.border = borde
            if celda.row % 2 == 0:
                celda.fill = gris
            if celda.column in idx_moneda:
                celda.number_format = '"₡"#,##0'

    # Ancho de columnas según el contenido
    for col in ws.columns:
        largo = max(len(str(c.value)) if c.value is not None else 0 for c in col)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(largo + 4, 40)

    ws.freeze_panes = "A2"             # encabezado siempre visible
    ws.auto_filter.ref = ws.dimensions # filtros en el encabezado


# ---------------------------------------------------------------
# PASO 1: EXTRACT
# ---------------------------------------------------------------

# 1. Conexión
conn = sqlite3.connect(DB)

# 2a. Consulta con WHERE: productos caros
q_where = """
SELECT id_producto, nombre, categoria, precio
FROM productos
WHERE precio > 50
"""
df_caros = pd.read_sql_query(q_where, conn)

# 2b. Consulta con JOIN: cada venta con el nombre del cliente y del producto
q_join = """
SELECT v.id_venta, c.nombre AS cliente, c.ciudad,
       p.nombre AS producto, p.categoria, p.precio,
       v.cantidad, v.fecha
FROM ventas v
JOIN clientes c ON v.id_cliente = c.id_cliente
JOIN productos p ON v.id_producto = p.id_producto
"""
df_ventas = pd.read_sql_query(q_join, conn)

# 2c. Consulta con GROUP BY: cantidad de productos por categoría
q_group = """
SELECT categoria, COUNT(*) AS num_productos, AVG(precio) AS precio_promedio
FROM productos
GROUP BY categoria
"""
df_categorias = pd.read_sql_query(q_group, conn)

# Revisión rápida
for nombre, d in [("caros", df_caros), ("ventas", df_ventas), ("categorias", df_categorias)]:
    print(f"--- {nombre}: {d.shape[0]} filas ---")
    print(d.head(), "\n")

conn.close()

# ---------------------------------------------------------------
# PASO 2: DETECCIÓN DE PROBLEMAS DE CALIDAD
# ---------------------------------------------------------------

# revision de ventas duplicadas
cols = ["cliente", "producto", "cantidad", "fecha"]
print("Duplicados sin contar id_venta:", df_ventas.duplicated(subset=cols).sum())
print(df_ventas[df_ventas.duplicated(subset=cols, keep=False)].sort_values(cols))

# Revision de valores numericos sin sentido
print(df_ventas["cantidad"].describe())
print(df_ventas["cantidad"].value_counts().sort_index())
print(df_ventas[df_ventas["cantidad"] <= 0])

# ---------------------------------------------------------------
# PASO 3: TRANSFORM
# ---------------------------------------------------------------

df = df_ventas.copy()
print("Filas al inicio:", len(df))

# 1. Ciudad: normalizar formato (quita espacios, primera letra de cada palabra en mayúscula)
df["ciudad"] = df["ciudad"].str.strip().str.title()

# Los vacíos: se marcan como "Sin dato" en vez de borrar la venta
df["ciudad"] = df["ciudad"].fillna("Sin dato")
print(df["ciudad"].value_counts())

# 2. Duplicados: quitar la copia y conservar el original
df = df.drop_duplicates(subset=cols, keep="first")
print("Filas tras quitar duplicados:", len(df))

# 3. Cantidades negativas: eliminar esas filas
df = df[df["cantidad"] > 0]
print("Filas tras quitar cantidades negativas:", len(df))

# Total de cada venta
df["total"] = df["cantidad"] * df["precio"]
print(df[["id_venta", "producto", "cantidad", "precio", "total"]].head())
print("Total general:", df["total"].sum())

# Reporte agregado: ventas totales por categoría
reporte = (
    df.groupby("categoria", as_index=False)["total"]
      .sum()
      .sort_values("total", ascending=False)
)
print(reporte)

# Gráfico del reporte (se guarda ahora, se muestra al final)
fig, ax = plt.subplots(figsize=(7, 4))
ax.bar(reporte["categoria"], reporte["total"])
ax.set_title("Ventas totales por categoría")
ax.set_xlabel("Categoría")
ax.set_ylabel("Total vendido (₡)")
fig.tight_layout()
fig.savefig(ruta_png)

# ---------------------------------------------------------------
# PASO 4: LOAD (archivos nuevos en la carpeta de esta ejecución)
# ---------------------------------------------------------------

try:
    reporte.to_csv(ruta_reporte, index=False, encoding="utf-8-sig")
    df.to_csv(ruta_limpias, index=False, encoding="utf-8-sig")

    # Excel con dos hojas: el reporte y las ventas limpias
    with pd.ExcelWriter(ruta_excel, engine="openpyxl") as writer:
        reporte.to_excel(writer, sheet_name="Reporte", index=False)
        df.to_excel(writer, sheet_name="Ventas limpias", index=False)

        ws_rep = writer.sheets["Reporte"]
        ws_ven = writer.sheets["Ventas limpias"]

        dar_formato(ws_rep, columnas_moneda=("total",))
        dar_formato(ws_ven, columnas_moneda=("precio", "total"))

        # Fila de total al final del reporte
        ultima = len(reporte) + 1
        ws_rep.cell(row=ultima + 1, column=1, value="Total").font = Font(bold=True)
        celda_total = ws_rep.cell(row=ultima + 1, column=2, value=f"=SUM(B2:B{ultima})")
        celda_total.font = Font(bold=True)
        celda_total.number_format = '"₡"#,##0'

        # Gráfico nativo de Excel junto al reporte
        grafico = BarChart()
        grafico.title = "Ventas totales por categoría"
        grafico.y_axis.title = "Total vendido (₡)"
        grafico.x_axis.title = "Categoría"
        grafico.add_data(Reference(ws_rep, min_col=2, min_row=1, max_row=ultima), titles_from_data=True)
        grafico.set_categories(Reference(ws_rep, min_col=1, min_row=2, max_row=ultima))
        grafico.legend = None
        grafico.width, grafico.height = 16, 8
        ws_rep.add_chart(grafico, "D2")

    print("Archivos generados en:", CARPETA)
    for archivo in sorted(CARPETA.iterdir()):
        print(" -", archivo.name)
finally:
    # Se muestra el gráfico pase lo que pase con el guardado
    plt.show()
    plt.close(fig)