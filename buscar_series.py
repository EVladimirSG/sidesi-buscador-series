import streamlit as st
import pandas as pd
import io
import re
from time import sleep

# =========================================================
# ⚙️ CONFIGURACIÓN GENERAL
# =========================================================
st.set_page_config(page_title="SISTEMA DE BUSCAR SERIES, VS", page_icon="📘", layout="centered")

st.title("📘 **BUSCADOR DE SERIES PARA DESCARGO DE MATERIALES**")
st.write("""
Esta herramienta es para busca las series instaladas en los centro de cierre de cada mes
         en mantenimiento DTH PROINTEL, o garantias.
""")

# =========================================================
# 📂 CARGA DE ARCHIVOS DESDE LA INTERFAZ
# =========================================================
archivo_series = st.file_uploader("📂 Subir archivo con series o existencias del centro(.xlsx)", type=["xlsx"])
archivo_cierres = st.file_uploader("📂 Subir archivo de CIERRES (.xlsx)", type=["xlsx"])

if not archivo_series or not archivo_cierres:
    st.warning("Por favor, suba ambos archivos (.xlsx) antes de iniciar la búsqueda.")
    st.stop()

try:
    st.info("Cargando archivos... Esto puede tardar unos segundos...")
    sleep(0.8)

    # ✅ Leer los archivos desde la interfaz de Streamlit
    df_series = pd.read_excel(archivo_series)
    df_cierres = pd.read_excel(archivo_cierres)

    # Normalizar nombres de columnas
    df_series.columns = df_series.columns.str.strip().str.upper()
    df_cierres.columns = df_cierres.columns.str.strip().str.upper()

    # Mostrar vista previa
    with st.expander("Vista previa del archivo de SERIES/EXISTENCIAS"):
        st.dataframe(df_series.head(5), use_container_width=True)
    with st.expander("Vista previa del archivo de CIERRES"):
        st.dataframe(df_cierres.head(5), use_container_width=True)

    # Buscar columnas que contengan "SERIE"
    col_series = [c for c in df_series.columns if "SERIE" in c]
    if not col_series:
        st.error("No se encontró ninguna columna con 'SERIE' en el archivo de series.")
        st.stop()

    # Combinar todas las columnas de series
    todas_series = []
    for col in col_series:
        todas_series += df_series[col].dropna().astype(str).tolist()

    todas_series = list(set([s.strip() for s in todas_series if s.strip()]))
    total_series = len(todas_series)
    st.info(f"🔍 Buscando coincidencias de {total_series} series en los comentarios...")

    # =========================================================
    # 🔎 BÚSQUEDA DE SERIES
    # =========================================================
    progreso = st.progress(0)
    coincidencias = []
    no_encontradas = []

    for i, serie in enumerate(todas_series):
        serie_str = str(serie).strip().lower()
        encontrado = False
        for _, fila in df_cierres.iterrows():
            texto = " ".join(fila.astype(str)).lower()
            if re.search(re.escape(serie_str), texto):
                fila_copy = fila.copy()
                fila_copy["SERIE"] = serie  # agrega la serie encontrada
                coincidencias.append(fila_copy)
                encontrado = True
                break
        if not encontrado:
            no_encontradas.append(serie)
        progreso.progress(int(((i + 1) / total_series) * 100))

    # =========================================================
    # RESULTADOS
    # =========================================================
    st.success("Búsqueda completada con éxito.")
    st.write(f"**Series encontradas:** {len(coincidencias)}")
    st.write(f"**Series no encontradas:** {len(no_encontradas)}")

    if coincidencias:
        df_resultados = pd.DataFrame(coincidencias)

        # --- Crear columnas faltantes si no existen ---
        columnas_finales = ["SERIE", "ND", "ZONA", "SEMANA", "ABRV_UI", "DEP", "COM_REP", "DD_TECI", "F_REP"]
        for col in columnas_finales:
            if col not in df_resultados.columns:
                df_resultados[col] = ""

        # --- Filtrar y reordenar las columnas ---
        df_resultados = df_resultados[columnas_finales]

        # --- Normalizar la columna ZONA ---
        df_resultados["ZONA"] = (
            df_resultados["ZONA"]
            .astype(str)
            .str.upper()
            .replace({
                "URBANO": "CENTRO",
                "RURAL": "ORIENTE",
                "OCCIDENTAL": "OCCIDENTE",
                "PARACENTRAL": "PARACENTRAL"
            })
        )

        # --- Mostrar resultados ---
        st.dataframe(df_resultados.head(30), use_container_width=True)

        # --- Exportar resultados a Excel ---
        fecha_actual = pd.Timestamp.now().strftime("%Y-%m-%d_%H-%M")
        nombre_archivo = f"Resultados_Buscador_SIDESI_{fecha_actual}.xlsx"

        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
            df_resultados.to_excel(writer, index=False, sheet_name="Coincidencias")
            if no_encontradas:
                pd.DataFrame(no_encontradas, columns=["SERIES_NO_ENCONTRADAS"]).to_excel(
                    writer, index=False, sheet_name="No_encontradas"
                )

            # Ajustar ancho automático de columnas
            for col_num, col_name in enumerate(df_resultados.columns):
                ancho = max(df_resultados[col_name].astype(str).map(len).max(), len(col_name)) + 2
                writer.sheets["Coincidencias"].set_column(col_num, col_num, ancho)

        st.download_button(
            label="Descargar resultados en Excel",
            data=buffer.getvalue(),
            file_name=nombre_archivo,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    else:
        st.warning("No se encontraron coincidencias.")

    # Mostrar series no encontradas
    if no_encontradas:
        with st.expander("Ver series no encontradas"):
            st.write(no_encontradas)

except Exception as e:
    st.error(f"🚨Error durante el procesamiento: {e}")







