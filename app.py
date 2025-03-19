import streamlit as st
import pandas as pd
from rapidfuzz import process, fuzz  # 🚀 Más rápido que FuzzyWuzzy
from openpyxl import Workbook
from io import BytesIO
from PIL import Image
import streamlit as st

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

            # 📊 **Estadísticas**
            total_base1 = len(df1_sin_dup)
            total_base2 = len(df2_sin_dup)
            coincidencias = len(df_coincidencias)
            sin_coincidencia = len(df_sin_coincidencia)
            porcentaje1 = f"{(coincidencias / total_base1 * 100):.2f}%" if total_base1 > 0 else "0.00%"
            porcentaje2 = f"{(coincidencias / total_base2 * 100):.2f}%" if total_base2 > 0 else "0.00%"

            df_estadisticas = pd.DataFrame({
                "Métrica": ["Total registros", "Coincidencias", "Sin coincidencia", "Porcentaje coincidencia"],
                f"Base {col1}": [total_base1, coincidencias, sin_coincidencia, porcentaje1],
                f"Base {col2}": [total_base2, coincidencias, sin_coincidencia, porcentaje2]
            })

            # 🖥️ **Mostrar Resultados**
            st.write("### 📊 Coincidencias Encontradas")
            st.data_editor(df_coincidencias, num_rows="dynamic")

            st.write("### ❌ Registros Sin Coincidencia")
            st.data_editor(df_sin_coincidencia, num_rows="dynamic")

            st.write("### 📈 Estadísticas")
            st.dataframe(df_estadisticas)

            # 📥 **Función para Descargar Reporte en Excel**
            @st.cache_data
            def convertir_a_excel(df_coincidencias, df_sin_coincidencia, df3, duplicados_dict):
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_coincidencias.to_excel(writer, sheet_name="Coincidencias", index=False)
                    df_sin_coincidencia.to_excel(writer, sheet_name="Sin Coincidencia", index=False)
                    df3.to_excel(writer, sheet_name="Estadísticas", index=False)
                    
                    # Agregar hojas separadas para cada columna con duplicados
                    for nombre, df in duplicados_dict.items():
                        df.to_excel(writer, sheet_name=nombre[:31], index=False)  # Limita a 31 caracteres
                    
                return output.getvalue()

            excel_data = convertir_a_excel(df_coincidencias, df_sin_coincidencia, df_estadisticas, duplicados_dict)

            st.download_button(
                label="📥 Descargar reporte en Excel",
                data=excel_data,
                file_name="reporte-GoXperts.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
