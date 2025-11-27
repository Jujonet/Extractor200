import fitz  # PyMuPDF
import pandas as pd
import re
import math
from io import BytesIO

CASILLA_INI = 1501
CASILLA_FIN = 2493

NUMERO_EURO = re.compile(r'-?\d{1,3}(?:\.\d{3})*,\d{2}$')

def extraer_casillas(pdf_bytes: bytes,
                     casilla_ini: int = CASILLA_INI,
                     casilla_fin: int = CASILLA_FIN,
                     y_tol: float = 2.5) -> pd.DataFrame:
    """
    Devuelve un DataFrame con columnas 'Casilla' y 'Valor'
    para todas las casillas del rango [casilla_ini, casilla_fin].
    Si no hay importe visible en el PDF, deja el valor en blanco.
    """
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")

    casillas_rango = [f"{i:05d}" for i in range(casilla_ini, casilla_fin + 1)]
    casillas_encontradas: dict[str, str] = {}

    for page in doc:
        words = page.get_text("words")  # x0, y0, x1, y1, text, block, line, word

        # Candidatos a importes (tienen coma decimal)
        valores = []
        casillas = []

        for x0, y0, x1, y1, texto, *_ in words:
            t = texto.strip()

            # Importes estilo "17.772,60" o "-1.000,00"
            if NUMERO_EURO.fullmatch(t):
                valores.append((x0, y0, x1, y1, t))
                continue

            # Códigos de casilla de 5 dígitos dentro del rango
            if re.fullmatch(r'\d{5}', t):
                n = int(t)
                if casilla_ini <= n <= casilla_fin:
                    casillas.append((x0, y0, x1, y1, t))

        if not casillas or not valores:
            continue

        # Ordenamos los importes por coordenada vertical para buscar el más cercano
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

                # Solo miramos importes que están a la derecha de la casilla
                dx = vx0 - x1
                if dx < -2:
                    continue

                dist = math.hypot(dx, dy)
                if mejor_dist is None or dist < mejor_dist:
                    mejor_dist = dist
                    mejor_valor = vtexto

            if mejor_valor is not None:
                casillas_encontradas[cod] = mejor_valor

    # Construimos el DataFrame completo 01501–02493, rellenando blancos
    filas = [
        {"Casilla": c, "Valor": casillas_encontradas.get(c, "")}
        for c in casillas_rango
    ]
    return pd.DataFrame(filas)
