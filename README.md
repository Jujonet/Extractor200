# Extractor2002
Extractor del modelo 200 de la AEAT
**Demo interna CBNK · Extracción automática de casillas del Impuesto de Sociedades**

---

## ¿Qué es esto?

Una app web sencilla que coge el PDF oficial del **Modelo 200** que te pasa el cliente y te saca los importes de todas las casillas en un Excel. Sin copiar a mano. Sin errores de transcripción. Sin sufrir.

El rango que barre por defecto es el **00100–02600**, que cubre holgadamente lo que aparece en cualquier Modelo 200 actual.

---

## Cómo se usa

1. Abre la app en el navegador.
2. Sube el PDF del Modelo 200 (el que exporta directamente la AEAT, sin proteger).
3. Espera dos segundos mientras la máquina hace lo que tú no quieres hacer.
4. Revisa en pantalla que las casillas clave tienen el importe correcto.
5. Descarga el Excel. Nombre automático: igual que el PDF pero con extensión `.xlsx`.

---

## Qué genera

Un Excel con dos columnas:

| Casilla | Valor      |
|---------|------------|
| 00500   | 17772.60   |
| 00501   | 0.00       |
| 00502   |            |
| ...     | ...        |

- Las casillas que aparecen en el PDF llevan su importe como número real.
- Las que no existen en ese PDF se quedan en blanco (no explota, tranquilo).

---

## Requisitos técnicos

```
streamlit
pymupdf        # El import es fitz. Sí. No preguntes.
pandas
xlsxwriter
```

Instalar todo de golpe:
```bash
pip install streamlit pymupdf pandas xlsxwriter
```

Arrancar la app:
```bash
streamlit run app.py
```

---

## Limitaciones conocidas (o sea, lo que puede fallar)

- **PDFs escaneados**: si el PDF es una imagen (un escaneo del papel físico), la extracción devuelve vacío. Necesita texto nativo, es decir, el PDF generado por la AEAT o exportado desde su sede electrónica.
- **PDFs protegidos**: si el cliente te manda el PDF con contraseña o restricciones de copia, tampoco funciona. Pídele el original sin proteger.
- **Maquetación rara**: si Hacienda algún día decide cambiar el diseño del formulario de forma drástica, puede que alguna casilla no empareje bien con su importe. Revisión manual recomendada para las casillas más críticas.

---

## Aviso legal (el clásico)

Versión de prueba de concepto. Validar siempre los importes clave antes de usarlo en análisis de riesgo real. El código no miente, pero la maquetación del PDF de Hacienda a veces tiene sus cosas.

