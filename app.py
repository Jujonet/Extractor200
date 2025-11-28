import re
import math
import os
import base64
from io import BytesIO

import fitz          # PyMuPDF (sí, el import es "fitz", no me preguntes por qué que bastante jodido estoy yo ya...)
import pandas as pd
import streamlit as st


# --------------------------------------------
# CONFIGURACIÓN DEL INVENTO
# --------------------------------------------

# Rango de casillas que queremos barrer.
# Como enteros, luego ya los convertimos a "01000", "01001", etc.
CASILLA_INICIO = 1000    # En el PDF que manejo arranca en la 1501 pero le metemos la 1000 no vaya a ser el demonio...
CASILLA_FIN = 2600       # En el PDF que manejo solo está hasta la 2493 pero metemos hasta 2600 para estar preparados para el futuro

# Patrón para reconocer importes tipo:
#   1.234,56   -1.234,56   0,00   12.345.678,90
# Si el Modelo 200 cambia, venimos aquí a renegociar con Hacienda.
PATRON_IMPORTE = re.compile(r"-?\d{1,3}(?:\.\d{3})*,\d{2}$")

def logo_svg(path: str, width: int = 160) -> str:
    """
    Devuelve un <img> HTML con un SVG embebido en base64.
    Traducido: logo vectorial que no se ve como el culo al escalarlo.
    """
    with open(path, "rb") as f:
        svg_bytes = f.read()
    b64 = base64.b64encode(svg_bytes).decode("utf-8")
    return f'<img src="data:image/svg+xml;base64,{b64}" width="{width}">'


# --------------------------------------------
# EXTRACCIÓN DE CASILLAS DESDE EL PDF
# --------------------------------------------

def extraer_casillas(pdf_bytes: bytes, y_tol: float = 2.5) -> pd.DataFrame:
    """
    Lee el PDF con PyMuPDF y devuelve un DataFrame con columnas:
      - Casilla: códigos de 5 dígitos entre CASILLA_INICIO y CASILLA_FIN
      - Valor: texto tal cual impreso ("17.772,60", "0,00", "")

    pdf_bytes: contenido del PDF en bruto.
    y_tol: tolerancia vertical para decidir si casilla e importe están alineados.
           Si lo subo mucho empezará a emparejar cosas que no tocan y la liamos.
    """

    # Abrimos el PDF desde bytes.
    # Esto lo vi en StackOverflow y como funciona, entra en la categoría "no se toca".
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

        # Distancia euclídea: básicamente mide lo cerca que está el importe
        # de la casilla en la página. Podría usar algo más simple, pero ya que 
        # existe math.hypot() y queda de ingeniero friki que te cagas, lo uso.
        # Preciso, limpio y además sube el ego técnico.
                
                dist = math.hypot(dx, dy)

                if mejor_dist is None or dist < mejor_dist:
                    mejor_dist = dist
                    mejor_valor = vtexto

            if mejor_valor is not None:
                casillas_encontradas[cod] = mejor_valor

    # Construimos la tabla completa CASILLA_INICIO–CASILLA_FIN,
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
#
# Esta función existe por una razón muy simple:
#
# El PDF trae los números en “formato español”, es decir:
#   - miles con puntos   →  1.234.567
#   - decimales con coma →  89,45
#
# Python, en cambio, habla anglosajon:
#   - miles sin separador → 1234567
#   - decimales con punto → 89.45
#
# Si intento convertir directamente “1.234,56” con float(), Python
# me mandaba al carallo (“ValueError”) porque no lo entiende.
#
# Así que:
#   1) quitamos los puntos de miles
#   2) cambiamos la coma por punto
#   3) convertimos a float
#   4) Rezamos
#
# Ejemplos:
#   "17.772,60" → "17772.60" → 17772.60
#   "0,00"      → "0.00"     → 0.0
#   ""          → None       (celda vacía en Excel)
#
# Si algo viene roto o raro, devuelvo None para no estrellarme....
#

def str_eu_a_float(x: str):
    """
    "17.772,60" -> 17772.60 (float).
    ""          -> None (para dejar celda en blanco).

    Si algo se rompe, devolvemos None y arreando....
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
        # Si llega algo rarísimo, mejor lo dejamos como texto y luego ya si eso.
        return None


# --------------------------------------------
# APP STREAMLIT (INTERFAZ JUJO + CBNK CORPORATE) 
# AQUI ME GUSTÉ A MÍ MISMO
# --------------------------------------------

def main():
    """
    Interfaz principal de la app Streamlit.
    Versión Jujo: explica lo justo, que no hace falta más...
    """

   
    st.set_page_config(
        page_title="Extractor Modelo 200",
        page_icon="📄",
        layout="wide",
    )

    # CSS para esconder menú y footer de Streamlit y dejarlo más "chulo"
    HIDE_STREAMLIT_STYLE = """
        <style>
        #MainMenu {visibility: hidden;}
        footer {visibility: hidden;}
        header {visibility: visible;}
        </style>
    """
    st.markdown(HIDE_STREAMLIT_STYLE, unsafe_allow_html=True)

    # Texto de rango generado a partir de las constantes, así si lo cambio no se queda viejo
    rango_texto = f"{CASILLA_INICIO:05d}–{CASILLA_FIN:05d}"

    # --- Cabecera limpia ---
    st.markdown(
        f"""
        <h1 style="margin-bottom:0.2rem;">Extractor Modelo 200</h1>
        <p style="color:#666; margin-top:0;">
        Demo interna · Relleno automático de casillas {rango_texto} a partir del PDF del impuesto.
        Aunque no existan todas esas casillas en el PDF, las intenta bajar igualmente.
        </p>
        """,
        unsafe_allow_html=True,
    )

    # --- Sidebar para comerciales ---
    with st.sidebar:
        # Logo corporativo en SVG, pero bien hecho, no la mierda de PNG que me amargaba
        try:
            html_logo = logo_svg("static/cbnk-logo.svg", width=160)
            st.markdown(html_logo, unsafe_allow_html=True)
        except Exception:
            pass

   
        st.markdown("### ¿Qué hace esto?")
        st.write(
            f"- Lee el PDF oficial del Modelo 200.\n"
            f"- Busca los importes de las casillas {rango_texto}.\n"
            "- Genera un Excel listo para copiar/pegar en tu hoja de trabajo."
        )

        st.markdown("### Cómo usarlo")
        st.write(
            "1. Coge cada PDF del Modelo 200 que te pasa el cliente (sin proteger).\n"
            "2. Súbelo aquí.\n"
            "3. Revisa que las casillas clave tengan el importe correcto.\n"
            "4. Descarga el Excel."
        )

        st.markdown("---")
        st.markdown(
            "<small>Versión de prueba de concepto. "
            "Es una simple prueba de cómo podría ser. No me hago responsable de nada.</small>",
            unsafe_allow_html=True,
        )

    # --- Subida de fichero ---
    uploaded = st.file_uploader("📎 Sube el PDF del Modelo 200", type=["pdf"])

    if not uploaded:
        st.info("Sube un PDF para empezar. Idealmente el PDF exportado directamente de la AEAT que te dio el cliente.")
        return

    # Capturo nombre del fichero PDF para usarlo como base del .xlsx
    pdf_filename = uploaded.name or "modelo200.pdf"
    base_name, _ = os.path.splitext(pdf_filename)
    excel_filename = f"{base_name}.xlsx"

    pdf_bytes = uploaded.getvalue()

    with st.spinner("Procesando PDF…"):
        df = extraer_casillas(pdf_bytes)

    st.success("PDF procesado. Revisa las casillas antes de descargar el Excel no vaya a ser el demonio....")

    # --- Vista previa en pantalla ---
    st.subheader("Vista previa de casillas y valores")
    st.dataframe(df, use_container_width=True, height=600)

    # --- Generación del Excel ---
    buffer = BytesIO()

    # Aquí paso de "modo pandas" a "modo yo controlo":
    # escribo la hoja a mano con XlsxWriter para asegurarme de que
    # el formato numérico se aplica bien.... Y aún así no queda al 100% pero yo ya paso...
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        workbook = writer.book

        # Creo la hoja yo mismo (no uso df.to_excel).
        worksheet = workbook.add_worksheet("Casillas")
        writer.sheets["Casillas"] = worksheet

        # Formatos básicos
        formato_header = workbook.add_format({"bold": True})
        # Formato numérico tipo ##.###,##  → en Excel: "#.##0,00" 
        # Pero por alguna p*ta razón excel pasa de mi.... pero no rompe nada....
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

            # Columna B: intento convertir a número y aplicar formato
            num = str_eu_a_float(valor_str)
            if num is not None:
                # Número real con formato ##.###,##
                # Segundo intento y segundo fracaso
                worksheet.write_number(i, 1, num, formato_num)
            else:
                # Si no hay valor o no se puede convertir, escribimos texto o dejamos en blanco
                if isinstance(valor_str, str) and valor_str.strip():
                    worksheet.write_string(i, 1, valor_str)
                else:
                    worksheet.write_blank(i, 1, None)

        # Ajusto anchos de columnas para que quede bonito....
        worksheet.set_column(0, 0, 8)   # Casilla
        worksheet.set_column(1, 1, 18)  # Valor

    buffer.seek(0)

    # Botón de descarga con icono chulo 
    st.download_button(
        label="📥 Descargar Excel",
        data=buffer,
        file_name=excel_filename,
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
    )

    # Pie de página minimalista
    st.markdown(
        "<div style='color:#999; font-size:0.75rem; margin-top:1rem;'>"
        "Generado automáticamente a partir del PDF del Modelo 200. "
        "Validar siempre los importes clave antes de usarlo, que hay partes cogidas con alfileres..."
        "</div>",
        unsafe_allow_html=True,
    )


if __name__ == "__main__":
    main()
