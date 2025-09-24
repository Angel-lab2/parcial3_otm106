import os
import pandas as pd
import matplotlib.pyplot as plt

# 1. Crear carpeta "datos" si no existe
if not os.path.exists("datos"):
    os.makedirs("datos")

# 2. Ruta del archivo Excel
file_path = os.path.join("datos", "ejemplo.xlsx")

# 3. Datos ficticios para las hojas
data = {
    "Barras1": {
        "Producto": ["A", "B", "C", "D", "E"],
        "Ventas": [50, 70, 40, 90, 60]
    },
    "Barras2": {
        "Día": ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes"],
        "Alumnos": [30, 28, 32, 29, 31]
    },
    "Barras3": {
        "Semana": [1, 2, 3, 4, 5],
        "Producción": [100, 120, 95, 110, 130]
    },
    "Pastel1": {
        "Empresa": ["X", "Y", "Z", "W", "V"],
        "% Mercado": [25, 30, 20, 15, 10]
    },
    "Pastel2": {
        "Categoría": ["Alimentos", "Transporte", "Vivienda", "Educación", "Ocio"],
        "Gasto": [400, 150, 600, 200, 100]
    },
    "Pastel3": {
        "Actividad": ["Trabajo", "Estudio", "Descanso", "Deporte", "Ocio"],
        "Horas": [8, 4, 6, 2, 4]
    }
}

# 4. Crear el archivo Excel con las 6 hojas
with pd.ExcelWriter(file_path, engine="xlsxwriter") as writer:
    for sheet, values in data.items():
        df = pd.DataFrame(values)
        df.to_excel(writer, sheet_name=sheet, index=False)

print(f"Archivo Excel creado en: {file_path}")

# 5. Crear gráficos con matplotlib y guardarlos como imágenes
for sheet, values in data.items():
    df = pd.DataFrame(values)

    if "Barras" in sheet:
        plt.figure(figsize=(6, 4))
        plt.bar(df.iloc[:, 0], df.iloc[:, 1], color="skyblue")
        plt.title(f"Gráfico {sheet}")
        plt.xlabel(df.columns[0])
        plt.ylabel(df.columns[1])
        plt.savefig(os.path.join("datos", f"{sheet}.png"))
        plt.close()

    elif "Pastel" in sheet:
        plt.figure(figsize=(6, 4))
        plt.pie(df.iloc[:, 1], labels=df.iloc[:, 0], autopct="%1.1f%%")
        plt.title(f"Gráfico {sheet}")
        plt.savefig(os.path.join("datos", f"{sheet}.png"))
        plt.close()

print(" Gráficos guardados en la carpeta 'datos'")

