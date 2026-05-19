import streamlit as st
import pandas as pd
from rapidfuzz import fuzz
from io import BytesIO
import re
import unicodedata
import plotly.graph_objects as go
import time

st.set_page_config(
    page_title="Matching PRO MAX",
    layout="wide",
    page_icon="🧠",
    initial_sidebar_state="collapsed"
)

# -------------------------
# CSS PERSONALIZADO
# -------------------------
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@400;600;700&family=JetBrains+Mono:wght@400;600&display=swap');

html, body, [class*="css"] {
    font-family: 'Space Grotesk', sans-serif;
}

.stApp {
    background: #f8fafc;
    min-height: 100vh;
}

h1 {
    color: #1e293b !important;
    font-size: 2.6rem !important;
    font-weight: 700 !important;
    letter-spacing: -1px;
}

h2, h3 {
    color: #334155 !important;
    font-weight: 600 !important;
}

.score-badge {
    display: inline-block;
    font-family: 'JetBrains Mono', monospace;
    font-size: 3.5rem;
    font-weight: 700;
    padding: 1rem 2.5rem;
    border-radius: 20px;
    margin: 1rem 0;
}

.score-healthy { background: #f0fdf4; color: #16a34a; border: 2px solid #16a34a; }
.score-warning { background: #fffbeb; color: #d97706; border: 2px solid #d97706; }
.score-critical { background: #fef2f2; color: #dc2626; border: 2px solid #dc2626; }

.section-header {
    background: #f1f5f9;
    border-left: 4px solid #6366f1;
    padding: 0.6rem 1rem;
    border-radius: 0 8px 8px 0;
    margin: 1.5rem 0 1rem 0;
    color: #1e293b;
    font-size: 1.1rem;
    font-weight: 600;
}

.stDataFrame {
    border-radius: 12px !important;
    overflow: hidden !important;
}

.stButton > button {
    background: linear-gradient(135deg, #6366f1, #3b82f6) !important;
    color: white !important;
    border: none !important;
    border-radius: 10px !important;
    font-weight: 600 !important;
    padding: 0.5rem 2rem !important;
    transition: opacity 0.2s ease !important;
}

.stButton > button:hover {
    opacity: 0.85 !important;
}

.warning-box {
    background: #fffbeb;
    border: 1px solid #fbbf24;
    border-radius: 10px;
    padding: 0.8rem 1.2rem;
    color: #92400e;
    margin: 0.5rem 0;
    font-size: 0.9rem;
}

div[data-testid="stMetricValue"] {
    font-family: 'JetBrains Mono', monospace !important;
    font-size: 1.8rem !important;
}
</style>
""", unsafe_allow_html=True)

# -------------------------
# HEADER
# -------------------------
st.markdown("# 🧠 Matching PRO MAX")
st.markdown("<p style='color:#64748b; margin-top:-10px; margin-bottom:2rem;'>Inteligencia de datos · Matching fuzzy · Calidad de bases</p>", unsafe_allow_html=True)

# -------------------------
# LIMPIEZA
# -------------------------
STOPWORDS = {"sa", "s.a", "ltda", "inc", "corp", "company", "co", "srl", "s.a.s", "sas", "de", "la", "el", "los", "las"}

@st.cache_data(show_spinner=False)
def limpiar_texto(texto):
    texto = str(texto).lower().strip()
    texto = ''.join(
        c for c in unicodedata.normalize('NFD', texto)
        if unicodedata.category(c) != 'Mn'
    )
    texto = re.sub(r'[^a-z0-9\s]', '', texto)
    palabras = [p for p in texto.split() if p not in STOPWORDS]
    return " ".join(palabras)

def limpiar_serie(serie):
    return serie.dropna().apply(limpiar_texto)

# -------------------------
# SCORE
# -------------------------
def calcular_score(a, b):
    score = (
        fuzz.token_set_ratio(a, b) * 0.5 +
        fuzz.partial_ratio(a, b) * 0.3 +
        fuzz.token_sort_ratio(a, b) * 0.2
    )
    if a in b or b in a:
        score += 10
    return min(100, score)

# -------------------------
# MATCHING GLOBAL
# -------------------------
@st.cache_data(show_spinner=False)
def matching_ultra(base1_tuple, base2_tuple):
    base1 = list(base1_tuple)
    base2 = list(base2_tuple)

    resultados = []
    usados_a = set()
    usados_b = set()

    # 1. MATCH EXACTO
    exactos_b = {b: j for j, b in enumerate(base2)}
    for i, a in enumerate(base1):
        if a in exactos_b:
            j = exactos_b[a]
            if j not in usados_b:
                resultados.append({"Base 1": a, "Base 2": base2[j], "Score": 100, "Estado": "✅ MATCH"})
                usados_a.add(i)
                usados_b.add(j)

    # 2. POSIBLES MATCHES (fuzzy)
    posibles = []
    for i, a in enumerate(base1):
        if i in usados_a:
            continue
        for j, b in enumerate(base2):
            if j in usados_b:
                continue
            score = calcular_score(a, b)
            if score >= 65:
                posibles.append((i, j, score))

    posibles.sort(key=lambda x: x[2], reverse=True)

    for i, j, score in posibles:
        if i in usados_a or j in usados_b:
            continue
        estado = "✅ MATCH" if score >= 85 else "⚠️ REVISAR"
        resultados.append({
            "Base 1": base1[i],
            "Base 2": base2[j],
            "Score": round(score, 2),
            "Estado": estado
        })
        usados_a.add(i)
        usados_b.add(j)

    # 3. NO MATCH BASE 1
    for i, a in enumerate(base1):
        if i not in usados_a:
            resultados.append({"Base 1": a, "Base 2": None, "Score": 0, "Estado": "❌ NO MATCH"})

    # 4. SOBRANTES BASE 2
    for j, b in enumerate(base2):
        if j not in usados_b:
            resultados.append({"Base 1": None, "Base 2": b, "Score": 0, "Estado": "❌ SOBRANTE B"})

    return pd.DataFrame(resultados)

# -------------------------
# DUPLICADOS
# -------------------------
@st.cache_data(show_spinner=False)
def analizar_duplicados(serie_tuple, nombre):
    serie = pd.Series(serie_tuple)
    conteo = serie.value_counts()
    duplicados = conteo[conteo > 1]
    df_dup = duplicados.reset_index()
    df_dup.columns = [nombre, "Cantidad"]
    total = len(serie)
    total_dup = int(duplicados.sum())
    pct = (total_dup / total * 100) if total > 0 else 0
    return df_dup, total_dup, round(pct, 2)

# -------------------------
# SIMILARES (con límite de seguridad)
# -------------------------
MAX_SIMILARES = 500

@st.cache_data(show_spinner=False)
def duplicados_similares(base_tuple, threshold=90):
    lista = list(base_tuple)
    n = len(lista)

    if n > MAX_SIMILARES:
        return pd.DataFrame(), True  # demasiado grande

    similares = []
    for i in range(n):
        for j in range(i + 1, n):
            score = fuzz.ratio(lista[i], lista[j])
            if score >= threshold and lista[i] != lista[j]:
                similares.append((lista[i], lista[j], score))

    return pd.DataFrame(similares, columns=["Valor 1", "Valor 2", "Similitud"]), False

# -------------------------
# KPI CALIDAD
# -------------------------
def calcular_calidad(df_match, total_base, total_dup1, total_dup2):
    match = len(df_match[df_match["Estado"] == "✅ MATCH"])
    revisar = len(df_match[df_match["Estado"] == "⚠️ REVISAR"])
    no_match = len(df_match[df_match["Estado"] == "❌ NO MATCH"])

    pct_match = (match / total_base * 100) if total_base > 0 else 0
    pct_revisar = (revisar / total_base * 100) if total_base > 0 else 0
    pct_no_match = (no_match / total_base * 100) if total_base > 0 else 0
    total_dup = total_dup1 + total_dup2
    pct_dup = (total_dup / (total_base * 2) * 100) if total_base > 0 else 0

    score = (
        pct_match * 1.0 +
        pct_revisar * 0.5 -
        pct_no_match * 0.7 -
        pct_dup * 0.5
    )
    score = max(0, min(100, score))

    return {
        "score": round(score, 2),
        "match": round(pct_match, 2),
        "revisar": round(pct_revisar, 2),
        "no_match": round(pct_no_match, 2),
        "duplicados": round(pct_dup, 2),
        "total_match": match,
        "total_revisar": revisar,
        "total_no_match": no_match,
    }

# -------------------------
# GRÁFICO PLOTLY
# -------------------------
def grafico_kpi(kpi):
    labels = ["✅ Match", "⚠️ Revisar", "❌ No Match", "🚨 Duplicados"]
    values = [kpi["match"], kpi["revisar"], kpi["no_match"], kpi["duplicados"]]
    colors = ["#34d399", "#fbbf24", "#ef4444", "#f97316"]

    fig = go.Figure()

    # Donut
    fig.add_trace(go.Pie(
        labels=labels,
        values=values,
        hole=0.65,
        marker=dict(colors=colors, line=dict(color='rgba(0,0,0,0)', width=0)),
        textinfo="label+percent",
        textfont=dict(family="Space Grotesk", size=13, color="white"),
        hovertemplate="%{label}: %{value:.1f}%<extra></extra>",
    ))

    fig.add_annotation(
        text=f"<b>{kpi['score']}</b><br><span style='font-size:12px'>/ 100</span>",
        x=0.5, y=0.5,
        font=dict(size=28, color="#1e293b", family="JetBrains Mono"),
        showarrow=False,
        xref="paper", yref="paper"
    )

    fig.update_layout(
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        showlegend=True,
        legend=dict(
            font=dict(color="#334155", family="Space Grotesk"),
            bgcolor="rgba(0,0,0,0)"
        ),
        margin=dict(t=20, b=20, l=20, r=20),
        height=320,
    )

    return fig

# -------------------------
# RENDER KPI SECTION
# -------------------------
def render_kpi(df_res, base1, base2, total_dup1, total_dup2):
    kpi = calcular_calidad(df_res, len(base1), total_dup1, total_dup2)

    st.markdown('<div class="section-header">📊 Score de Calidad de Base</div>', unsafe_allow_html=True)

    col_graf, col_nums = st.columns([1, 1])

    with col_graf:
        fig = grafico_kpi(kpi)
        st.plotly_chart(fig, use_container_width=True)

    with col_nums:
        score_class = "score-healthy" if kpi["score"] >= 80 else ("score-warning" if kpi["score"] >= 60 else "score-critical")
        label = "🟢 Base saludable" if kpi["score"] >= 80 else ("🟡 Base con problemas" if kpi["score"] >= 60 else "🔴 Base crítica")

        st.markdown(f"""
        <div style='margin-top:1.5rem'>
            <div class='score-badge {score_class}'>{kpi['score']} / 100</div>
            <p style='color:#94a3b8; margin-top:0.5rem; font-size:1rem'>{label}</p>
        </div>
        """, unsafe_allow_html=True)

        m1, m2 = st.columns(2)
        m3, m4 = st.columns(2)

        m1.metric("✅ Match", f"{kpi['match']}%", f"{kpi['total_match']} registros")
        m2.metric("⚠️ Revisar", f"{kpi['revisar']}%", f"{kpi['total_revisar']} registros")
        m3.metric("❌ No Match", f"{kpi['no_match']}%", f"{kpi['total_no_match']} registros")
        m4.metric("🚨 Duplicados", f"{kpi['duplicados']}%")

    return kpi

# -------------------------
# RENDER DUPLICADOS
# -------------------------
def render_duplicados(base1, base2, col1_name, col2_name):
    st.markdown('<div class="section-header">🔍 Calidad de Datos — Duplicados</div>', unsafe_allow_html=True)

    dup1, total1, pct1 = analizar_duplicados(tuple(base1), col1_name)
    dup2, total2, pct2 = analizar_duplicados(tuple(base2), col2_name)

    c1, c2 = st.columns(2)
    with c1:
        if total1 > 0:
            st.error(f"**{col1_name}**: {total1} duplicados ({pct1}%)")
            st.dataframe(dup1, use_container_width=True, hide_index=True)
        else:
            st.success(f"**{col1_name}**: sin duplicados ✓")

    with c2:
        if total2 > 0:
            st.error(f"**{col2_name}**: {total2} duplicados ({pct2}%)")
            st.dataframe(dup2, use_container_width=True, hide_index=True)
        else:
            st.success(f"**{col2_name}**: sin duplicados ✓")

    return total1, total2

# -------------------------
# RENDER SIMILARES
# -------------------------
def render_similares(base, nombre):
    st.markdown(f'<div class="section-header">⚠️ Posibles Errores Humanos — {nombre}</div>', unsafe_allow_html=True)

    n = len(base)
    if n > MAX_SIMILARES:
        st.markdown(f"""
        <div class="warning-box">
            ⚡ <b>Análisis omitido:</b> {nombre} tiene {n} registros (límite: {MAX_SIMILARES}).
            El análisis de similares en bases grandes puede tardar varios minutos.
            Filtra o reduce la base para activarlo.
        </div>
        """, unsafe_allow_html=True)
        return pd.DataFrame()

    similares, _ = duplicados_similares(tuple(base))
    if not similares.empty:
        st.dataframe(similares.sort_values("Similitud", ascending=False), use_container_width=True, hide_index=True)
    else:
        st.success("Sin errores tipográficos similares detectados ✓")
    return similares

# -------------------------
# EXPORT EXCEL
# -------------------------
def exportar_excel(df_res, dup1, dup2, similares1, similares2, kpi_dict):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_res.to_excel(writer, sheet_name="Matching", index=False)
        if not dup1.empty:
            dup1.to_excel(writer, sheet_name="Duplicados Col1", index=False)
        if not dup2.empty:
            dup2.to_excel(writer, sheet_name="Duplicados Col2", index=False)
        if not similares1.empty:
            similares1.to_excel(writer, sheet_name="Similares Col1", index=False)
        if not similares2.empty:
            similares2.to_excel(writer, sheet_name="Similares Col2", index=False)
        pd.DataFrame([kpi_dict]).to_excel(writer, sheet_name="KPI", index=False)
    return output.getvalue()

# =========================================
# UI PRINCIPAL
# =========================================
modo = st.radio(
    "Selecciona modo de análisis",
    ["📄 Mismo archivo", "📂 Multi archivo"],
    horizontal=True
)

st.markdown("---")

# =========================
# MODO: MISMO ARCHIVO
# =========================
if modo == "📄 Mismo archivo":

    archivo = st.file_uploader("📎 Sube tu archivo Excel", type=["xlsx"])

    if archivo:
        excel = pd.ExcelFile(archivo)
        hoja = st.selectbox("📋 Selecciona la hoja", excel.sheet_names)
        df = pd.read_excel(excel, sheet_name=hoja)

        cols_obj = df.select_dtypes(include="object").columns.tolist()
        if len(cols_obj) < 2:
            st.error("El archivo necesita al menos 2 columnas de texto.")
            st.stop()

        col_a, col_b = st.columns(2)
        with col_a:
            col1 = st.selectbox("Columna 1 (Base maestra)", cols_obj, key="c1")
        with col_b:
            col2 = st.selectbox("Columna 2 (A comparar)", cols_obj, index=min(1, len(cols_obj)-1), key="c2")

        if col1 == col2:
            st.warning("⚠️ Selecciona dos columnas diferentes.")
            st.stop()

        if st.button("🚀 Ejecutar Análisis"):
            with st.spinner("Limpiando y procesando datos..."):
                base1 = limpiar_serie(df[col1])
                base2 = limpiar_serie(df[col2])

            # DUPLICADOS
            total_dup1, total_dup2 = render_duplicados(base1, base2, col1, col2)
            dup1_df, _, _ = analizar_duplicados(tuple(base1), col1)
            dup2_df, _, _ = analizar_duplicados(tuple(base2), col2)

            # SIMILARES (ambas columnas)
            similares1 = render_similares(base1, col1)
            similares2 = render_similares(base2, col2)

            # MATCHING
            st.markdown('<div class="section-header">🔗 Resultado del Matching</div>', unsafe_allow_html=True)
            with st.spinner("Calculando matches..."):
                progress = st.progress(0, text="Analizando...")
                for i in range(0, 80, 20):
                    time.sleep(0.1)
                    progress.progress(i, text="Analizando...")
                df_res = matching_ultra(tuple(base1), tuple(base2))
                progress.progress(100, text="¡Listo!")
                time.sleep(0.3)
                progress.empty()

            st.dataframe(df_res, use_container_width=True, hide_index=True)

            # KPI
            kpi = render_kpi(df_res, base1, base2, total_dup1, total_dup2)

            # EXPORT
            st.markdown("---")
            excel_bytes = exportar_excel(df_res, dup1_df, dup2_df, similares1, similares2, kpi)
            st.download_button(
                "📥 Descargar Reporte Completo (.xlsx)",
                excel_bytes,
                "reporte_matching_pro.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

# =========================
# MODO: MULTI ARCHIVO
# =========================
if modo == "📂 Multi archivo":

    archivos = st.file_uploader(
        "📎 Sube múltiples archivos Excel",
        type=["xlsx"],
        accept_multiple_files=True
    )

    if archivos:
        if len(archivos) < 2:
            st.warning("⚠️ Sube al menos 2 archivos para comparar.")
            st.stop()

        nombres = [f.name for f in archivos]
        maestro = st.selectbox("⭐ Archivo maestro (Base 1)", nombres)

        hojas = {}
        columnas = {}

        st.markdown("**Configuración por archivo:**")
        for archivo in archivos:
            with st.expander(f"📄 {archivo.name}", expanded=True):
                excel = pd.ExcelFile(archivo)
                hoja = st.selectbox(f"Hoja", excel.sheet_names, key=f"hoja_{archivo.name}")
                df_temp = pd.read_excel(excel, sheet_name=hoja)
                cols = df_temp.select_dtypes(include="object").columns.tolist()
                col = st.selectbox(f"Columna a comparar", cols, key=f"col_{archivo.name}")
                hojas[archivo.name] = hoja
                columnas[archivo.name] = col

        if st.button("🚀 Ejecutar Matching Multi"):
            archivo_maestro = next(f for f in archivos if f.name == maestro)
            df_master = pd.read_excel(archivo_maestro, sheet_name=hojas[maestro])
            base1 = limpiar_serie(df_master[columnas[maestro]])

            resultados = []
            progress = st.progress(0, text="Procesando archivos...")
            total = len([f for f in archivos if f.name != maestro])

            for idx, archivo in enumerate([f for f in archivos if f.name != maestro]):
                df = pd.read_excel(archivo, sheet_name=hojas[archivo.name])
                base2 = limpiar_serie(df[columnas[archivo.name]])

                df_res = matching_ultra(tuple(base1), tuple(base2))
                df_res["Archivo"] = archivo.name
                resultados.append(df_res)

                progress.progress(int((idx + 1) / total * 100), text=f"Procesado: {archivo.name}")

            progress.empty()
            df_final = pd.concat(resultados, ignore_index=True)

            st.markdown('<div class="section-header">🔗 Resultado del Matching Global</div>', unsafe_allow_html=True)
            st.dataframe(df_final, use_container_width=True, hide_index=True)

            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_final.to_excel(writer, index=False, sheet_name="Matching Global")

            st.download_button(
                "📥 Descargar Resultado (.xlsx)",
                output.getvalue(),
                "matching_multi.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

# -------------------------
# FOOTER
# -------------------------
st.markdown("""
<div style='text-align:center; color:#94a3b8; margin-top:4rem; font-size:0.8rem; padding-bottom:2rem;'>
    Matching PRO MAX · Powered by RapidFuzz + Streamlit · v2.0
</div>
""", unsafe_allow_html=True)
