import streamlit as st
import pandas as pd
from rapidfuzz import process, fuzz  # 🚀 Más rápido que FuzzyWuzzy
from openpyxl import Workbook
from io import BytesIO
from PIL import Image

# 📌 Cargar la imagen original
imagen = Image.open("go-xpert.png")

# 📌 Redimensionar la imagen al 50% de su tamaño original
nuevo_tamano = (imagen.width // 2, imagen.height // 2)
imagen_reducida = imagen.resize(nuevo_tamano)

# 📌 Mostrar la imagen en Streamlit (sin advertencias)
st.image(imagen_reducida)

# 🔍 **Título de la aplicación**
st.title("🔍 Analizador de Coincidencias - SMART")

st.write("Sube dos archivos de Excel y selecciona las hojas y columnas a comparar.")

# 📂 **Carga de archivos**
archivo1 = st.file_uploader("📂 Sube el primer archivo Excel", type=["xlsx"])
archivo2 = st.file_uploader("📂 Sube el segundo archivo Excel", type=["xlsx"])

if archivo1 and archivo2:
    excel1 = pd.ExcelFile(archivo1)
    excel2 = pd.ExcelFile(archivo2)

    hoja1 = st.selectbox("📑 Selecciona la hoja del primer archivo", excel1.sheet_names)
    hoja2 = st.selectbox("📑 Selecciona la hoja del segundo archivo", excel2.sheet_names)

    if hoja1 and hoja2:
        df1 = pd.read_excel(excel1, sheet_name=hoja1)
        df2 = pd.read_excel(excel2, sheet_name=hoja2)

        col1 = st.multiselect("📊 Selecciona las columnas del primer archivo", df1.columns)
        col2 = st.multiselect("📊 Selecciona las columnas del segundo archivo", df2.columns)

        if col1 and col2 and len(col1) == len(col2):
            umbral = st.slider("🎯 Umbral de similitud (0-100)", min_value=0, max_value=100, value=80)

            def limpiar_texto(df, columnas):
                for col in columnas:
                    df[col] = df[col].astype(str).str.lower().str.strip()
                    df[col] = df[col].str.replace(r'\s+', ' ', regex=True)
                return df

            df1 = limpiar_texto(df1, col1)
            df2 = limpiar_texto(df2, col2)

            def emparejar_bases(df1, df2, col1, col2, threshold):
                emparejados = []
                base1_set = df1[col1].astype(str).apply(tuple, axis=1).tolist()
                base2_set = df2[col2].astype(str).apply(tuple, axis=1).tolist()
                progreso = st.progress(0)

                total = len(base1_set)
                for i, row_tuple in enumerate(base1_set):
                    progreso.progress((i + 1) / total)
                    match = process.extractOne(row_tuple, base2_set, scorer=fuzz.token_sort_ratio)

                    if match and match[1] >= threshold:
                        emparejados.append(list(row_tuple) + list(match[0]) + [match[1], 'Coincidencia'])
                        base2_set.remove(match[0])
                    else:
                        emparejados.append(list(row_tuple) + [None] * len(col2) + [0, 'Sin coincidencia'])

                for row in base2_set:
                    if any(row):
                        emparejados.append([None] * len(col1) + list(row) + [0, 'Sin coincidencia'])

                progreso.empty()
                return pd.DataFrame(emparejados, columns=col1 + col2 + ['Similitud (%)', 'Estado'])

            df_emparejados = emparejar_bases(df1, df2, col1, col2, umbral)
            df_coincidencias = df_emparejados[df_emparejados["Estado"] == "Coincidencia"]
            df_sin_coincidencia = df_emparejados[df_emparejados["Estado"] == "Sin coincidencia"]

            df_estadisticas = pd.DataFrame({
                "Métrica": ["Total registros", "Coincidencias", "Sin coincidencia", "Porcentaje coincidencia"],
                f"Base {col1}": [len(df1), len(df_coincidencias), len(df_sin_coincidencia), f"{(len(df_coincidencias) / len(df1) * 100):.2f}%"],
                f"Base {col2}": [len(df2), len(df_coincidencias), len(df_sin_coincidencia), f"{(len(df_coincidencias) / len(df2) * 100):.2f}%"]
            })

            @st.cache_data
            def convertir_a_excel():
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_emparejados.to_excel(writer, sheet_name="Emparejamiento", index=False)
                    df_coincidencias.to_excel(writer, sheet_name="Coincidencias", index=False)
                    df_sin_coincidencia.to_excel(writer, sheet_name="Sin Coincidencia", index=False)
                    df_estadisticas.to_excel(writer, sheet_name="Estadísticas", index=False)
                return output.getvalue()

            excel_data = convertir_a_excel()
            st.download_button("📥 Descargar reporte en Excel", data=excel_data, file_name="reporte-GoXperts.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
