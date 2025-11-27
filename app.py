import re
import math
import os
from io import BytesIO

import fitz  # PyMuPDF (sí, el import es "fitz", porque alguien lo decidió así)
import pandas as pd
import streamlit as st

# --------------------------------------------
# CONFIGURACIÓN DEL EXPERIMENTO
# --------------------------------------------

# Rango de casillas que queremos barrer.
# Como enteros, luego ya los convertimos a "01501", "01502", etc.
CASILLA_INICIO = 1501
CASILLA_FIN = 3600 
# Aunque el PDF sólo tiene hasta 03400 metemos 3600 por si en el futuro se añaden más....

# Patrón para reconocer importes tipo:
#   1.234,56   -1.234,56   0,00   12.345.678,90
# Si el Modelo 200 cambia, venimos aquí y renegociamos con Hacienda.
PATRON_IMPORTE = re.compile(r"-?\d{1,3}(?:\.\d{3})*,\d{2}$")


# --------------------------------------------
# EXTRACCIÓN DE CASILLAS DESDE EL PDF
# --------------------------------------------

def extraer_casillas(pdf_bytes: bytes, y_tol: float = 2.5) -> pd.DataFrame:
    """
    Lee el PDF con PyMuPDF y devuelve un DataFrame con columnas:
      - Casilla: "01501", "01502", ...
      - Valor: texto tal cual impreso ("17.772,60", "0,00", "")

    pdf_bytes: contenido del PDF en bruto.
    y_tol: tolerancia vertical para decidir si casilla e importe están alineados.
           Si lo subes mucho empezará a emparejar números que no tocan.
    """

    # Abrimos el PDF desde bytes.
    # Esto lo vi en StackOverflow; entra en categoría "funciona → no se toca".
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")

    # Aquí guardaremos solo las casillas que SÍ encuentren un importe.
    casillas_encontradas: dict[str, str] = {}

    # Recorremos todas las páginas del PDF.
    for page in doc:
        # get_text("words") devuelve una lista de:
        # (x0, y0, x1, y1, text, block_no, line_no, word_no)
        words = page.get_text("words")

        valores = []   # candidatos a importes
        casillas = []  # candidatos a códigos de casilla

        for x0, y0, x1, y1, texto, *_ in words:
            t = texto.strip()

            # 1) ¿Tiene pinta de importe tipo "17.772,60"?
            if PATRON_IMPORTE.fullmatch(t):
                valores.append((x0, y0, x1, y1, t))
                continue

            # 2) ¿Tiene pinta de casilla de 5 dígitos dentro del rango?
            if re.fullmatch(r"\d{5}", t):
                n = int(t)
                if CASILLA_INICIO <= n <= CASILLA_FIN:
                    casillas.append((x0, y0, x1, y1, t))

        # Si en la página no hay ni casillas ni importes, pasamos.
        if not casillas or not valores:
            continue

        # Ordenamos los importes por altura vertical (centro Y).
        # Ayuda a que el emparejamiento sea menos random.
        valores.sort(key=lambda w: (w[1] + w[3]) / 2)

        # Emparejamos cada casilla con el importe "más razonable":
        # alineado en vertical y a la derecha.
        for x0, y0, x1, y1, cod in casillas:
            y_c = (y0 + y1) / 2  # centro vertical de la casilla

            mejor_valor = None
            mejor_dist = None

            for vx0, vy0, vx1, vy1, vtexto in valores:
                vy_c = (vy0 + vy1) / 2
                dy = abs(vy_c - y_c)
                if dy > y_tol:
                    continue  # demasiado alto/bajo

                dx = vx0 - x1      # solo importes a la derecha
                if dx < -2:
                    continue

                # Distancia euclídea, porque si podemos meter un hypot,
                # lo metemos (matemáticas + ego).
                dist = math.hypot(dx, dy)

                if mejor_dist is None or dist < mejor_dist:
                    mejor_dist = dist
                    mejor_valor = vtexto

            if mejor_valor is not None:
                casillas_encontradas[cod] = mejor_valor

    # Construimos la tabla completa 01501–02493,
    # rellenando con "" cuando no hay importe.
    casillas_ordenadas = [f"{n:05d}" for n in range(CASILLA_INICIO, CASILLA_FIN + 1)]

    filas = []
    for c in casillas_ordenadas:
        filas.append(
            {
                "Casilla": c,
                # Si la casilla no tiene importe, dejamos cadena vacía.
                "Valor": casillas_encontradas.get(c, ""),
            }
        )

    return pd.DataFrame(filas)


# --------------------------------------------
# UTILIDAD: STRING EU -> FLOAT PYTHON
# --------------------------------------------

def str_eu_a_float(x: str):
    """
    "17.772,60" -> 17772.60 (float).
    ""          -> None (para dejar celda en blanco).

    Si algo se rompe, devolvemos None y ya.
    """
    if not isinstance(x, str):
        return None
    t = x.strip()
    if not t:
        return None
    # "17.772,60" -> "17772,60" -> "17772.60"
    t = t.replace(".", "").replace(",", ".")
    try:
        return float(t)
    except Exception:
        # Si llega algo rarísimo, mejor lo dejamos como texto luego.
        return None


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

    # Nombre del fichero PDF para usarlo como base del .xlsx
    pdf_filename = uploaded.name or "modelo200.pdf"
    base_name, _ = os.path.splitext(pdf_filename)
    excel_filename = f"{base_name}.xlsx"

    pdf_bytes = uploaded.getvalue()

    with st.spinner("Procesando PDF..."):
        df = extraer_casillas(pdf_bytes)

    # --- Vista en pantalla (tal cual texto del PDF) ---
    st.subheader("Vista previa de casillas y valores")
    st.dataframe(df, width="stretch", height=600)

    # --- Generación del Excel ---
    # Aquí pasamos de "modo pandas" a "modo yo controlo": escribimos la hoja
    # a mano con XlsxWriter para asegurarnos de que el formato se aplica.

    buffer = BytesIO()

    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        workbook = writer.book

        # Creamos la hoja nosotros mismos (no usamos df.to_excel)
        worksheet = workbook.add_worksheet("Casillas")
        writer.sheets["Casillas"] = worksheet

        # Formatos básicos
        formato_header = workbook.add_format({"bold": True})
        # Formato numérico tipo ##.###,##  → en Excel: "#.##0,00"
        formato_num = workbook.add_format({"num_format": "#.##0,00"})

        # Cabeceras
        worksheet.write(0, 0, "Casilla", formato_header)
        worksheet.write(0, 1, "Valor", formato_header)

        # Datos fila a fila
        for i, row in enumerate(df.itertuples(index=False), start=1):
            casilla = row.Casilla
            valor_str = row.Valor

            # Columna A: siempre texto
            worksheet.write_string(i, 0, casilla)

            # Columna B: intentamos convertir a número y aplicar formato
            num = str_eu_a_float(valor_str)
            if num is not None:
                # Número real con formato ##.###,##
                worksheet.write_number(i, 1, num, formato_num)
            else:
                # Si no hay valor o no se puede convertir, escribimos texto o dejamos en blanco
                if isinstance(valor_str, str) and valor_str.strip():
                    worksheet.write_string(i, 1, valor_str)
                else:
                    worksheet.write_blank(i, 1, None)

        # Ajustamos anchos de columnas
        worksheet.set_column(0, 0, 8)          # Casilla
        worksheet.set_column(1, 1, 18)         # Valor (formato ya aplicado por write_number)

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
