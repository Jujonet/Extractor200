import re
import math
from io import BytesIO

import fitz  # PyMuPDF (sí, el import es "fitz", porque razones)
import pandas as pd
import streamlit as st

# --------------------------------------------
# CONFIGURACIÓN BÁSICA DEL EXPERIMENTO
# --------------------------------------------

# Rango de casillas que queremos barrer.
# Como enteros, luego ya los formateamos a "01501", etc.
CASILLA_INICIO = 1501
CASILLA_FIN = 2493

# Patrón para reconocer importes tipo:
#   1.234,56   -1.234,56   0,00   12.345.678,90
# Si Hacienda cambia el formato, lloramos y lo actualizamos.
PATRON_IMPORTE = re.compile(r"-?\d{1,3}(?:\.\d{3})*,\d{2}$")


def extraer_casillas(pdf_bytes: bytes, y_tol: float = 2.5) -> pd.DataFrame:
    """
    Lee el PDF con PyMuPDF y devuelve un DataFrame con columnas:
      - Casilla: "01501", "01502", ...
      - Valor: cadena tal cual impresa ("17.772,60", "0,00", "")

    pdf_bytes: contenido del PDF en bruto.
    y_tol: tolerancia vertical para decidir si casilla e importe están alineados.
    """

    # Abrimos el PDF desde bytes. Esto lo vi en StackOverflow:
    # funciona, así que entra en la categoría "no tocar".
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")

    # Aquí guardaremos solo las casillas que SÍ encuentren un importe.
    casillas_encontradas: dict[str, str] = {}

    # Recorremos todas las páginas del PDF.
    for page in doc:
        # words: (x0, y0, x1, y1, text, block_no, line_no, word_no)
        words = page.get_text("words")

        valores = []   # candidatos a importes
        casillas = []  # candidatos a códigos de casilla

        for x0, y0, x1, y1, texto, *_ in words:
            t = texto.strip()

            # 1) ¿Tiene pinta de importe tipo "17.772,60"?
            if PATRON_IMPORTE.fullmatch(t):
                valores.append((x0, y0, x1, y1, t))
                continue

            # 2) ¿Tiene pinta de casilla de 5 dígitos en nuestro rango?
            if re.fullmatch(r"\d{5}", t):
                n = int(t)
                if CASILLA_INICIO <= n <= CASILLA_FIN:
                    casillas.append((x0, y0, x1, y1, t))

        # Si en la página no hay casillas o no hay importes, a otra cosa.
        if not casillas or not valores:
            continue

        # Ordenamos los importes por altura vertical (centro Y).
        valores.sort(key=lambda w: (w[1] + w[3]) / 2)

        # Emparejamos cada casilla con el importe "más razonable"
        # (alineado en vertical y a la derecha).
        for x0, y0, x1, y1, cod in casillas:
            y_c = (y0 + y1) / 2  # centro vertical de la casilla

            mejor_valor = None
            mejor_dist = None

            for vx0, vy0, vx1, vy1, vtexto in valores:
                vy_c = (vy0 + vy1) / 2
                dy = abs(vy_c - y_c)

                # Si verticalmente están muy lejos, descartamos.
                if dy > y_tol:
                    continue

                # Queremos importes a la DERECHA de la casilla.
                dx = vx0 - x1
                if dx < -2:
                    continue

                # Distancia euclídea, porque si podemos meter un hypot,
                # lo metemos (también cortesía de StackOverflow).
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
                # Si la casilla no tiene importe, devolvemos cadena vacía.
                "Valor": casillas_encontradas.get(c, ""),
            }
        )

    return pd.DataFrame(filas)


def convertir_a_float_eu(x: str) -> float | None:
    """
    Convierte un string tipo "17.772,60" a float Python 17772.60.

    - Si viene vacío, devolvemos None (celda vacía en Excel).
    - Si pasa algo raro, también devolvemos None (mejor vacío que mal).
    """
    if not x:
        return None

    # Quitamos puntos de miles y cambiamos coma por punto.
    #  "17.772,60" -> "17772,60" -> "17772.60"
    x_norm = x.replace(".", "").replace(",", ".")
    try:
        return float(x_norm)
    except Exception:
        # Si algo se escapa al patrón, no reventamos el mundo por una casilla.
        return None


def main():
    """
    Función principal de la app Streamlit.
    Si llegamos aquí sin petar por imports, ya vamos ganando.
    """

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

    # --- Vista en pantalla ---
    # Aquí mostramos el DataFrame tal cual lo hemos extraído:
    # valores como strings "17.772,60", que es lo que tú esperas ver.
    st.subheader("Vista previa de casillas y valores")
    st.dataframe(df, width="stretch", height=600)

    # --- Preparación para Excel ---
    # Para que Excel tenga la columna B como NÚMERO, hacemos una copia
    # del DataFrame y en esa copia convertimos "Valor" a float.
    df_excel = df.copy()
    df_excel["Valor"] = df_excel["Valor"].apply(convertir_a_float_eu)

    buffer = BytesIO()

    # Usamos XlsxWriter para poder aplicar formato numérico explícito.
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        # Volcamos df_excel: aquí "Valor" ya son floats / None
        df_excel.to_excel(writer, index=False, sheet_name="Casillas")

        # Jugamos un poco con el workbook/worksheet para formatear.
        workbook = writer.book
        worksheet = writer.sheets["Casillas"]

        # Formato numérico: miles con punto, decimales con coma.
        # Excel ya adaptará esto según la configuración regional del usuario.
        formato_num = workbook.add_format({"num_format": "#.##0,00"})

        # Columna B (segunda columna, cero-indexed -> col 1).
        # Le damos un ancho razonable y el formato numérico.
        worksheet.set_column(1, 1, 18, formato_num)

    buffer.seek(0)  # rebobinamos para que Streamlit pueda mandar el archivo

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
