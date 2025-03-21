import streamlit as st
 import pandas as pd
 from rapidfuzz import process, fuzz  # ⚡ MÁS RÁPIDO QUE FUZZYWUZZY
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
 
         col1 = st.selectbox("📊 Selecciona la columna del primer archivo", df1.columns)
         col2 = st.selectbox("📊 Selecciona la columna del segundo archivo", df2.columns)
 
         if col1 and col2:
             umbral = st.slider("🎯 Umbral de similitud (0-100)", min_value=0, max_value=100, value=80)
 
             base1 = df1[col1].dropna().astype(str).str.lower().str.strip().unique()
             base2 = df2[col2].dropna().astype(str).str.lower().str.strip().unique()
 
             # 🔍 **Detección de duplicados**
             duplicados_base1 = df1[df1.duplicated(subset=[col1], keep=False)][col1].unique()
             duplicados_base2 = df2[df2.duplicated(subset=[col2], keep=False)][col2].unique()
 
             # 🔄 **Optimización del Emparejamiento con rapidfuzz**
             def emparejar_bases(base1, base2, threshold):
                 emparejados = []
                 base2_usada = set()
                 progreso = st.progress(0)  # Indicador de progreso
 
                 total = len(base1)
                 for i, nombre1 in enumerate(base1):
                     progreso.progress((i + 1) / total)
 
                     if pd.isna(nombre1):
                         continue
 
                     match = process.extractOne(
                         nombre1, 
                         [n for n in base2 if n not in base2_usada], 
                         scorer=fuzz.token_sort_ratio
                     )
 
                     if match and match[1] >= threshold:
                         emparejados.append([nombre1, match[0], match[1], 'Coincidencia'])
                         base2_usada.add(match[0])
                     else:
                         emparejados.append([nombre1, None, 0, 'Sin coincidencia'])
 
                 # Agregar elementos no coincidentes de base2
                 for nombre2 in base2:
                     if nombre2 not in base2_usada:
                         emparejados.append([None, nombre2, 0, 'Sin coincidencia'])
 
                 progreso.empty()  # Oculta la barra de progreso
                 return pd.DataFrame(emparejados, columns=[f'Base {col1}', f'Base {col2}', 'Similitud (%)', 'Estado'])
 
             # 🔥 **Ejecutar emparejamiento**
             df_emparejados = emparejar_bases(base1, base2, umbral)
 
             # 📊 **Estadísticas**
             total_base1 = len(base1)
             total_base2 = len(base2)
             coincidencias = len(df_emparejados[df_emparejados["Estado"] == "Coincidencia"])
             porcentaje1 = f"{(coincidencias / total_base1 * 100):.2f}%" if total_base1 > 0 else "0.00%"
             porcentaje2 = f"{(coincidencias / total_base2 * 100):.2f}%" if total_base2 > 0 else "0.00%"
 
             df_estadisticas = pd.DataFrame({
                 "Métrica": ["Total registros", "Coincidencias", "Porcentaje coincidencia"],
                 f"Base {col1}": [total_base1, coincidencias, porcentaje1],
                 f"Base {col2}": [total_base2, coincidencias, porcentaje2]
             })
 
             # 🖥️ **Mostrar Resultados**
             st.write("### 📊 Resultados del Análisis")
             st.dataframe(df_emparejados)
 
             st.write("### 🔄 Duplicados encontrados")
             st.write(f"📌 **Duplicados en {col1}**:")
             st.dataframe(pd.DataFrame(duplicados_base1, columns=[col1]))
             st.write(f"📌 **Duplicados en {col2}**:")
             st.dataframe(pd.DataFrame(duplicados_base2, columns=[col2]))
 
             st.write("### 📈 Estadísticas")
             st.dataframe(df_estadisticas)
 
             # 📥 **Función para Descargar Reporte en Excel**
             @st.cache_data
             def convertir_a_excel(df1, df2, df3):
                 output = BytesIO()
                 with pd.ExcelWriter(output, engine='openpyxl') as writer:
                     df1.to_excel(writer, sheet_name="Emparejamiento", index=False)
                     pd.DataFrame(duplicados_base1, columns=[col1]).to_excel(writer, sheet_name=f"Duplicados {col1}", index=False)
                     pd.DataFrame(duplicados_base2, columns=[col2]).to_excel(writer, sheet_name=f"Duplicados {col2}", index=False)
                     df3.to_excel(writer, sheet_name="Estadísticas", index=False)
                 return output.getvalue()
 
             excel_data = convertir_a_excel(df_emparejados, duplicados_base1, df_estadisticas)
             st.download_button(
                 label="📥 Descargar reporte en Excel",
                 data=excel_data,
                 file_name="reporte-GoXperts.xlsx",
                 mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
             )
