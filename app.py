import io
import re

import pdfplumber
import pandas as pd
import streamlit as st


# --- Parámetros del problema ---
CASILLA_INICIO = 1501
CASILLA_FIN = 2493  # inclusivo


def extraer_texto_pdf(pdf_file) -> str:
    """Devuelve todo el texto del PDF concatenado."""
    with pdfplumber.open(pdf_file) as pdf:
        partes = []
        for page in pdf.pages:
            t = page.extract_text()
            if t:
                partes.append(t)
    return "\n".join(partes)


def buscar_valor_casilla(texto: str, casilla: str) -> str:
    """
    Heurística simple:
    Busca la casilla (ej. '01501') y coge un bloque numérico después.
    Esto habrá que ajustarlo con PDFs reales.
    """
    # Ejemplos que intenta capturar:
    # 01501  1.234,56
    # 01501: 1234,56
    # Casilla 01501  -123.456,78
    patron = rf"{casilla}[^\n\r0-9\-+]*([\-+]?\d[\d\.\,]*)"
    m = re.search(patron, texto)
    if m:
        return m.group(1).strip()
    return ""  # si no encuentra nada, se deja en blanco


def extraer_casillas(pdf_file):
    """Devuelve un dict {casilla: valor_str} para todo el rango 01501–02493."""
    texto = extraer_texto_pdf(pdf_file)
    resultados = {}
    for n in range(CASILLA_INICIO, CASILLA_FIN + 1):
        casilla = f"{n:05d}"
        valor = buscar_valor_casilla(texto, casilla)
        resultados[casilla] = valor
    return resultados


def construir_dataframe(casillas_dict):
    casillas_ordenadas = sorted(casillas_dict.keys())
    data = {
        "Casilla": casillas_ordenadas,
        "Valor": [casillas_dict[c] for c in casillas_ordenadas],
    }
    return pd.DataFrame(data)


def df_a_excel_bytes(df: pd.DataFrame) -> bytes:
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Casillas")
    buffer.seek(0)
    return buffer.getvalue()


# --- Interfaz Streamlit ---

st.set_page_config(page_title="Extractor Modelo 200", layout="wide")

st.title("Extractor Modelo 200 – Casillas 01501–02493")

st.write(
    "Sube un PDF del Modelo 200. "
    "Se buscarán las casillas de la 01501 a la 02493. "
    "Si una casilla no aparece o no tiene valor, se deja en blanco."
)

uploaded_file = st.file_uploader("Sube el PDF del Modelo 200", type=["pdf"])

if uploaded_file is not None:
    with st.spinner("Procesando PDF..."):
        casillas_dict = extraer_casillas(uploaded_file)
        df = construir_dataframe(casillas_dict)

    st.subheader("Vista previa de casillas y valores")
    st.dataframe(df, use_container_width=True, height=600)

    excel_bytes = df_a_excel_bytes(df)

    st.download_button(
        label="📥 Descargar Excel",
        data=excel_bytes,
        file_name="modelo200_casillas_01501_02493.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
else:
    st.info("Esperando que subas un PDF…")
