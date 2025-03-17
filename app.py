import streamlit as st
import pandas as pd
from rapidfuzz import process, fuzz  # 🚀 Más rápido que FuzzyWuzzy
from openpyxl import Workbook
from io import BytesIO

st.set_page_config(page_title="Analizador de Excel", page_icon="🔍", layout="wide")
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

            # 🔍 **Detección de duplicados**
            duplicados_base1 = df1[df1.duplicated(subset=col1, keep=False)][col1].drop_duplicates()
            duplicados_base2 = df2[df2.duplicated(subset=col2, keep=False)][col2].drop_duplicates()

            # 🔥 **Optimización del Emparejamiento**
            def emparejar_bases(df1, df2, col1, col2, threshold):
                emparejados = []
                
                # ✅ CORRECCIÓN: Convertir las filas en tuplas correctamente
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

                # Agregar elementos no coincidentes de base2
                for row in base2_set:
                    emparejados.append([None] * len(col1) + list(row) + [0, 'Sin coincidencia'])

                progreso.empty()
                return pd.DataFrame(emparejados, columns=col1 + col2 + ['Similitud (%)', 'Estado'])

            # 🔥 **Ejecutar emparejamiento**
            df_emparejados = emparejar_bases(df1, df2, col1, col2, umbral)

            # 📊 **Filtrado de Resultados**
            filtro_min_similitud = st.slider("📊 Filtrar por porcentaje mínimo de coincidencia", 0, 100, 50)
            df_filtrado = df_emparejados[df_emparejados["Similitud (%)"] >= filtro_min_similitud]

            # 📊 **Estadísticas**
            total_base1 = len(df1)
            total_base2 = len(df2)
            coincidencias = len(df_filtrado[df_filtrado["Estado"] == "Coincidencia"])
            porcentaje1 = f"{(coincidencias / total_base1 * 100):.2f}%" if total_base1 > 0 else "0.00%"
            porcentaje2 = f"{(coincidencias / total_base2 * 100):.2f}%" if total_base2 > 0 else "0.00%"

            df_estadisticas = pd.DataFrame({
                "Métrica": ["Total registros", "Coincidencias", "Porcentaje coincidencia"],
                f"Base {col1}": [total_base1, coincidencias, porcentaje1],
                f"Base {col2}": [total_base2, coincidencias, porcentaje2]
            })

            # 🖥️ **Mostrar Resultados**
            st.write("### 📊 Resultados del Análisis")
            st.data_editor(df_filtrado, num_rows="dynamic")

            st.write("### 🔄 Duplicados encontrados")
            st.write(f"📌 **Duplicados en {col1}**:")
            st.dataframe(duplicados_base1)
            st.write(f"📌 **Duplicados en {col2}**:")
            st.dataframe(duplicados_base2)

            st.write("### 📈 Estadísticas")
            st.dataframe(df_estadisticas)

            # 📥 **Función para Descargar Reporte en Excel**
            @st.cache_data
            def convertir_a_excel(df1, df2, df3):
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df1.to_excel(writer, sheet_name="Emparejamiento", index=False)
                    df2.to_excel(writer, sheet_name="Duplicados", index=False)
                    df3.to_excel(writer, sheet_name="Estadísticas", index=False)
                return output.getvalue()

            excel_data = convertir_a_excel(df_filtrado, duplicados_base1, df_estadisticas)
            st.download_button(
                label="📥 Descargar reporte en Excel",
                data=excel_data,
                file_name="reporte-GoXperts.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
