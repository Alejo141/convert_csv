import streamlit as st
import pandas as pd
import io
import openpyxl
from datetime import datetime, date
from decimal import Decimal

st.set_page_config(page_title="Excel → CSV", page_icon="📊", layout="centered")

st.title("📊 Conversor Excel → CSV")
st.markdown("Carga un archivo Excel y descárgalo como CSV separado por coma, **sin modificar el formato de los datos**.")

# --- Funciones de utilidad ---

def numero_a_str(valor):
    """
    Convierte un número float a string preservando decimales reales.
    - Enteros representados como float (5.0) → '5'
    - Con decimales (3.14) → '3.14'
    - Sin notación científica (1e-5 → '0.00001')
    """
    try:
        d = Decimal(str(valor)).normalize()
        _, _, exp = d.as_tuple()
        if exp >= 0:
            return str(int(valor))
        return format(d, 'f')
    except Exception:
        return str(valor)


def celda_a_str(valor):
    """Convierte el valor de una celda a string, forzando fechas a dd-mm-yyyy y preservando decimales."""
    if valor is None:
        return ""
    if isinstance(valor, datetime):
        return valor.strftime("%d-%m-%Y")
    if isinstance(valor, date):
        return valor.strftime("%d-%m-%Y")
    if isinstance(valor, float):
        return numero_a_str(valor)
    if isinstance(valor, int):
        return str(valor)
    return str(valor)


def leer_hoja_sin_conversion(ws):
    """Lee una hoja de openpyxl celda a celda y devuelve lista de listas con valores como texto."""
    filas = []
    for fila in ws.iter_rows(values_only=True):
        filas.append([celda_a_str(celda) for celda in fila])
    return filas


def filas_a_csv_bytes(filas):
    """Convierte lista de listas a bytes CSV."""
    output = io.StringIO()
    for fila in filas:
        celdas_escapadas = []
        for celda in fila:
            if ',' in celda or '"' in celda or '\n' in celda:
                celda = '"' + celda.replace('"', '""') + '"'
            celdas_escapadas.append(celda)
        output.write(",".join(celdas_escapadas) + "\n")
    return output.getvalue().encode("utf-8-sig")


# --- UI ---

archivo = st.file_uploader("Selecciona un archivo Excel", type=["xlsx", "xls", "xlsm"])

if archivo:
    try:
        wb = openpyxl.load_workbook(archivo, data_only=True)
        hojas = wb.sheetnames

        hoja_sel = st.selectbox("Selecciona la hoja a convertir:", hojas)

        ws = wb[hoja_sel]

        filas = leer_hoja_sin_conversion(ws)
        n_filas = len(filas)
        n_cols = max((len(f) for f in filas), default=0)

        st.markdown(f"**Hoja:** `{hoja_sel}` · {n_filas} filas × {n_cols} columnas")

        if filas:
            preview_df = pd.DataFrame(filas[1:11], columns=filas[0] if filas else [])
            st.markdown("**Vista previa (primeras 10 filas):**")
            st.dataframe(preview_df, use_container_width=True)

        csv_bytes = filas_a_csv_bytes(filas)
        nombre_csv = archivo.name.rsplit(".", 1)[0] + f"_{hoja_sel}.csv"

        st.download_button(
            label="⬇️ Descargar CSV",
            data=csv_bytes,
            file_name=nombre_csv,
            mime="text/csv",
        )

        st.success(f"Archivo listo: **{nombre_csv}** ({len(csv_bytes):,} bytes)")

    except Exception as e:
        st.error(f"Error al procesar el archivo: {e}")

st.markdown("---")
st.caption("Fechas en dd-mm-yyyy · Decimales preservados · Enteros sin punto · Textos y vacíos sin transformación.")
