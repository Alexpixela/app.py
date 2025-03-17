import streamlit as st
import pandas as pd
from rapidfuzz import process, fuzz  # 🚀 Más rápido que FuzzyWuzzy
from openpyxl import Workbook
from io import BytesIO

# 📌 Estilos para centrar el banner
st.image(
    """
    <style>
        .banner-container {
            display: flex;
            justify-content: center;
            align-items: center;
            padding: 10px 0;
        }
        .banner-container img {
            max-width: 80%; /* Ajusta el tamaño del banner */
            height: auto;
        }
    </style>
    """,
    unsafe_allow_html=True
)

# 📌 Imagen del banner (Cambia por la URL o la ruta del archivo local)
st.image(
    """
    <div class="banner-container">
        <img src="go-xpert.png" alt="Banner de la compañía">
    </div>
    """,
    unsafe_allow_html=True
)

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

        # Selección de múltiples columnas
        col1 = st.multiselect("📊 Selecciona las columnas del primer archivo", df1.columns)
        col2 = st.multiselect("📊 Selecciona las columnas del segundo archivo", df2.columns)

        if col1 and col2 and len(col1) == len(col2):
            umbral = st.slider("🎯 Umbral de similitud (0-100)", min_value=0, max_value=100, value=80)

            # 🔄 **Normalización y limpieza de datos**
            def limpiar_texto(df, columnas):
                for col in columnas:
                    df[col] = df[col].astype(str).str.lower().str.strip()
                    df[col] = df[col].str.replace(r'\s+', ' ', regex=True)  # Quita espacios extra
                return df

            df1 = limpiar_texto(df1, col1)
            df2 = limpiar_texto(df2, col2)

            # 🔍 **Detección de duplicados por cada columna**
            duplicados_dict = {}
            for c1, c2 in zip(col1, col2):
                duplicados_dict[f"Duplicados {c1}"] = df1[df1.duplicated(subset=[c1], keep=False)][[c1]].drop_duplicates()
                duplicados_dict[f"Duplicados {c2}"] = df2[df2.duplicated(subset=[c2], keep=False)][[c2]].drop_duplicates()

            # 🔥 **Eliminamos los valores duplicados antes de comparar**
            df1_sin_dup = df1.drop_duplicates(subset=col1)
            df2_sin_dup = df2.drop_duplicates(subset=col2)

            # 🔥 **Optimización del Emparejamiento**
            def emparejar_bases(df1, df2, col1, col2, threshold):
                emparejados = []
                
                # ✅ Convertir filas en tuplas para comparación
                base1_set = df1[col1].astype(str).apply(tuple, axis=1).values.tolist()
                base2_set = df2[col2].astype(str).apply(tuple, axis=1).values.tolist()
                
                progreso = st.progress(0)

                total = len(base1_set)
                for i, row_tuple in enumerate(base1_set):
                    progreso.progress((i + 1) / total)

                    # Busca la mejor coincidencia en base2
                    match = process.extractOne(row_tuple, base2_set, scorer=fuzz.token_sort_ratio)

                    if match and match[1] >= threshold:
                        emparejados.append(list(row_tuple) + list(match[0]) + [match[1], 'Coincidencia'])
                        base2_set.remove(match[0])  # Evita reutilizar coincidencias
                    else:
                        emparejados.append(list(row_tuple) + [None] * len(col2) + [0, 'Sin coincidencia'])

                # Agregar elementos no coincidentes de base2, evitando celdas vacías
                for row in base2_set:
                    if any(row):  # Solo agrega filas si tienen algún valor válido
                        emparejados.append([None] * len(col1) + list(row) + [0, 'Sin coincidencia'])

                progreso.empty()
                return pd.DataFrame(emparejados, columns=col1 + col2 + ['Similitud (%)', 'Estado'])

            # 🔥 **Ejecutar emparejamiento sin duplicados**
            df_emparejados = emparejar_bases(df1_sin_dup, df2_sin_dup, col1, col2, umbral)

            # 📊 **Separar Coincidencias y Sin Coincidencia**
            df_coincidencias = df_emparejados[df_emparejados["Estado"] == "Coincidencia"]
            df_sin_coincidencia = df_emparejados[df_emparejados["Estado"] == "Sin coincidencia"]

            # 🖥️ **Mostrar Resultados**
            st.write("### 📊 Coincidencias Encontradas")
            st.data_editor(df_coincidencias, num_rows="dynamic")

            st.write("### ❌ Registros Sin Coincidencia")
            st.data_editor(df_sin_coincidencia, num_rows="dynamic")

