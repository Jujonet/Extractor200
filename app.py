import re
import math
from io import BytesIO

import fitz  # PyMuPDF
import pandas as pd
import streamlit as st

# Rango de casillas
CASILLA_INICIO = 1501
CASILLA_FIN = 2493

# Importes tipo 1.234,56 o -1.234,56
PATRON_IMPORTE = re.compile(r"-?\d{1,3}(?:\.\d{3})*,\d{2}$")


def extraer_casillas(pdf_bytes: bytes, y_tol: float = 2.5) -> pd.DataFrame:
    """
    Lee el PDF con PyMuPDF y devuelve un DataFrame con columnas:
    - Casilla (01501...02493)
    - Valor (importe encontrado o blanco)
    """
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    casillas_encontradas: dict[str, str] = {}

    for page in doc:
        # words: x0, y0, x1, y1, text, block_no, line_no, word_no
        words = page.get_text("words")

        valores = []
        casillas = []

        for x0, y0, x1, y1, texto, *_ in words:
            t = texto.strip()

            # Importes tipo "17.772,60" o "0,00"
            if PATRON_IMPORTE.fullmatch(t):
                valores.append((x0, y0, x1, y1, t))
                continue

            # Códigos de casilla de 5 dígitos dentro del rango
            if re.fullmatch(r"\d{5}", t):
                n = int(t)
                if CASILLA_INICIO <= n <= CASILLA_FIN:
                    casillas.append((x0, y0, x1, y1, t))

        if not casillas or not valores:
            continue

        # Ordenamos por altura para facilitar el emparejamiento
        valores.sort(key=lambda w: (w[1] + w[3]) / 2)

        for x0, y0, x1, y1, cod in casillas:
            y_c = (y0 + y1) / 2

            mejor_valor = None
            mejor_dist = None

            for vx0, vy0, vx1, vy1, vtexto in valores:
                vy_c = (vy0 + vy1) / 2
                dy = abs(vy_c - y_c)
                if dy > y_tol:
                    continue

                # Solo miramos importes a la derecha de la casilla
                dx = vx0 - x1
                if dx < -2:
                    continue

                dist = math.hypot(dx, dy)
                if mejor_dist is None or dist < mejor_dist:
                    mejor_dist = dist
                    mejor_valor = vtexto

            if mejor_valor is not None:
                casillas_encontradas[cod] = mejor_valor

    # Construimos el DataFrame completo 01501–02493, rellenando con blanco
    casillas_ordenadas = [f"{n:05d}" for n in range(CASILLA_INICIO, CASILLA_FIN + 1)]
    filas = [
        {"Casilla": c, "Valor": casillas_encontradas.get(c, "")}
        for c in casillas_ordenadas
    ]

    return pd.DataFrame(filas)


def main():
    st.set_page_config(page_title="Extractor Modelo 200", layout="wide")

    st.title("Extractor Modelo 200 – Casillas 01501–02493")
    st.write(
        "Sube un PDF del Modelo 200. "
        "Se buscan importes para las casillas 01501–02493. "
        "Si una casilla no tiene importe visible en el PDF, se deja en blanco."
    )

    uploaded = st.file_uploader("PDF del Modelo 200", type=["pdf"])

    if not uploaded:
        st.info("Sube un PDF para empezar.")
        return

    pdf_bytes = uploaded.getvalue()

    with st.spinner("Procesando PDF..."):
        df = extraer_casillas(pdf_bytes)

    st.subheader("Vista previa de casillas y valores")
    st.dataframe(df, width="stretch", height=600)

    buffer = BytesIO()
    df.to_excel(buffer, index=False, sheet_name="Casillas")
    buffer.seek(0)

    st.download_button(
        label="📥 Descargar Excel",
        data=buffer,
        file_name="modelo200_casillas_01501_02493.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
    )


if __name__ == "__main__":
    main()

