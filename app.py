import re
import math
import os
from io import BytesIO

import fitz  # PyMuPDF
import pandas as pd
import streamlit as st

# --------------------------------------------
# CONFIG
# --------------------------------------------

CASILLA_INICIO = 1501
CASILLA_FIN = 2493

# Importes tipo 1.234,56 / -1.234,56 / 0,00
PATRON_IMPORTE = re.compile(r"-?\d{1,3}(?:\.\d{3})*,\d{2}$")


# --------------------------------------------
# EXTRACCIÓN DESDE EL PDF
# --------------------------------------------

def extraer_casillas(pdf_bytes: bytes, y_tol: float = 2.5) -> pd.DataFrame:
    """
    Devuelve un DataFrame:
      Casilla: "01501"..."02493"
      Valor: cadena tal cual en el PDF ("17.772,60", "0,00" o "")
    """
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    casillas_encontradas: dict[str, str] = {}

    for page in doc:
        words = page.get_text("words")  # x0, y0, x1, y1, text, ...

        valores = []
        casillas = []

        for x0, y0, x1, y1, texto, *_ in words:
            t = texto.strip()

            # Candidatos a importe
            if PATRON_IMPORTE.fullmatch(t):
                valores.append((x0, y0, x1, y1, t))
                continue

            # Candidatos a código de casilla dentro del rango
            if re.fullmatch(r"\d{5}", t):
                n = int(t)
                if CASILLA_INICIO <= n <= CASILLA_FIN:
                    casillas.append((x0, y0, x1, y1, t))

        if not casillas or not valores:
            continue

        valores.sort(key=lambda w: (w[1] + w[3]) / 2)  # por altura

        for x0, y0, x1, y1, cod in casillas:
            y_c = (y0 + y1) / 2

            mejor_valor = None
            mejor_dist = None

            for vx0, vy0, vx1, vy1, vtexto in valores:
                vy_c = (vy0 + vy1) / 2
                dy = abs(vy_c - y_c)
                if dy > y_tol:
                    continue

                dx = vx0 - x1  # solo importes a la derecha
                if dx < -2:
                    continue

                dist = math.hypot(dx, dy)
                if mejor_dist is None or dist < mejor_dist:
                    mejor_dist = dist
                    mejor_valor = vtexto

            if mejor_valor is not None:
                casillas_encontradas[cod] = mejor_valor

    casillas_ordenadas = [f"{n:05d}" for n in range(CASILLA_INICIO, CASILLA_FIN + 1)]
    filas = [
        {
            "Casilla": c,
            "Valor": casillas_encontradas.get(c, ""),
        }
        for c in casillas_ordenadas
    ]
    return pd.DataFrame(filas)


# --------------------------------------------
# STRING EU -> FLOAT PYTHON
# --------------------------------------------

def str_eu_a_num(x: str):
    """'17.772,60' -> 17772.60 (float). '' -> NaN."""
    if not isinstance(x, str) or not x:
        return pd.NA
    x_norm = x.replace(".", "").replace(",", ".")
    try:
        return float(x_norm)
    except Exception:
        return pd.NA


# --------------------------------------------
# APP STREAMLIT
# --------------------------------------------

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

    pdf_filename = uploaded.name or "modelo200.pdf"
    base_name, _ = os.path.splitext(pdf_filename)
    excel_filename = f"{base_name}.xlsx"

    pdf_bytes = uploaded.getvalue()

    with st.spinner("Procesando PDF..."):
        df = extraer_casillas(pdf_bytes)

    # Vista previa: valores tal cual en texto
    st.subheader("Vista previa de casillas y valores")
    st.dataframe(df, width="stretch", height=600)

    # --- Preparación para Excel ---
    df_excel = df.copy()

    # 1) Pasamos "Valor" (string EU) -> floats reales
    serie_num = df_excel["Valor"].apply(str_eu_a_num)
    df_excel["Valor"] = pd.to_numeric(serie_num, errors="coerce")  # float64

    buffer = BytesIO()

    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        df_excel.to_excel(writer, index=False, sheet_name="Casillas")

        workbook = writer.book
        worksheet = writer.sheets["Casillas"]

        # Formato numérico tipo "##.###,##"
        # Excel lo entiende como "#.##0,00" y pone más puntos según haga falta.
        formato_num = workbook.add_format({"num_format": "#.##0,00"})

        # Aplicamos ancho + formato a la columna B (índice 1)
        worksheet.set_column(1, 1, 18, formato_num)

    buffer.seek(0)

    st.download_button(
        label="📥 Descargar Excel",
        data=buffer,
        file_name=excel_filename,
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
    )


if __name__ == "__main__":
    main()
