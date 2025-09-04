# streamlit_app.py
# Requisitos (requirements.txt sugerido):
# streamlit
# savReaderWriter
# openpyxl
# (opcional) pyreadstat  # solo si planeas migrar en el futuro

import collections
import collections.abc
collections.Iterable = collections.abc.Iterable  # Compat. savReaderWriter en Python 3.10+

import streamlit as st
from savReaderWriter import SavReader
from io import BytesIO
from pathlib import Path
from openpyxl import Workbook
import tempfile
from typing import Dict, Any, List, Tuple

# ======================
# Traducciones
# ======================
TEXTS = {
    "es": {
        "title": "Convertir SPSS a Excel",
        "subtitle": "Sube un archivo .sav/.zsav, aplica etiquetas y descarga Excel",
        "uploader": "📂 Arrastra o sube un archivo SPSS (.sav, .zsav)",
        "load_status": "Cargando archivo SPSS…",
        "success_load": "Archivo cargado correctamente ✅",
        "error_load": "Error al cargar archivo",
        "file_info": "**Archivo:** {name}  •  **Filas:** {rows:,}  •  **Columnas:** {cols}",
        "preview": "Vista previa de datos",
        "download": "💾 Descargar como Excel (.xlsx)",
        "saving": "Generando Excel…",
        "success_save": "Excel generado ✅",
        "toggle_lang": "🌐 English",
        "tips": "💡 Consejo: si tu archivo tiene muchas filas, la descarga puede tomar algunos segundos.",
        "no_file": "Sube un archivo para comenzar",
        "sheetname": "Datos"
    },
    "en": {
        "title": "Convert SPSS to Excel",
        "subtitle": "Upload a .sav/.zsav file, apply labels and download an Excel",
        "uploader": "📂 Drag & drop or upload an SPSS file (.sav, .zsav)",
        "load_status": "Loading SPSS file…",
        "success_load": "File loaded successfully ✅",
        "error_load": "Error loading file",
        "file_info": "**File:** {name}  •  **Rows:** {rows:,}  •  **Columns:** {cols}",
        "preview": "Data preview",
        "download": "💾 Download as Excel (.xlsx)",
        "saving": "Generating Excel…",
        "success_save": "Excel generated ✅",
        "toggle_lang": "🌐 Español",
        "tips": "💡 Tip: if your file is large, generating the download can take some seconds.",
        "no_file": "Upload a file to get started",
        "sheetname": "Data"
    },
}

# ======================
# Helpers principales
# ======================

def decode_bytes(x: Any) -> Any:
    if isinstance(x, bytes):
        return x.decode("utf-8", errors="ignore")
    return x

@st.cache_data(show_spinner=False)
def process_sav(file_bytes: bytes) -> Tuple[List[List[Any]], List[str], Dict[str, Dict[Any, str]]]:
    """Procesa el SAV con SavReader, aplica etiquetas de variables/valores y retorna (rows, headers, value_labels).
    - file_bytes: contenido del archivo .sav
    Retorna:
      rows: lista de filas con valores procesados y etiquetas aplicadas (sin la fila 0 de nombres)
      headers: lista de cabeceras finales (var_name (var_label) si existe, si no var_name)
      value_labels_dict: mapeo de etiquetas de valores por variable
    """
    # Escribir bytes a un archivo temporal para que SavReader pueda leerlo
    with tempfile.NamedTemporaryFile(delete=False, suffix=".sav") as tmp:
        tmp.write(file_bytes)
        tmp_path = tmp.name

    with SavReader(tmp_path, returnHeader=True, rawMode=False) as reader:
        header = reader.header
        varLabels = reader.varLabels
        valueLabels = reader.valueLabels
        records = reader.all()

    # Headers como diccionario index->var_name (str)
    headers_dict = {}
    for i, h in enumerate(header):
        headers_dict[i] = decode_bytes(h)

    # Etiquetas de variables
    var_labels_dict: Dict[str, str] = {}
    if varLabels:
        for var_key, var_label in varLabels.items():
            var_key_str = decode_bytes(var_key)
            var_labels_dict[var_key_str] = decode_bytes(var_label)

    # Etiquetas de valores
    value_labels_dict: Dict[str, Dict[Any, str]] = {}
    if valueLabels:
        for var_key, val_dict in valueLabels.items():
            var_key_str = decode_bytes(var_key)
            converted = {}
            for val_code, val_label in val_dict.items():
                converted[val_code] = decode_bytes(val_label)
            value_labels_dict[var_key_str] = converted

    # Headers finales (var_name (var_label))
    final_headers: List[str] = []
    for i in range(len(headers_dict)):
        var_name = headers_dict[i]
        label = var_labels_dict.get(var_name, "").strip()
        final_headers.append(f"{var_name} ({label})" if label else var_name)

    # Filas (omitimos la fila 0 si viniera con nombres)
    rows: List[List[Any]] = []
    for idx, record in enumerate(records):
        if idx == 0:
            # La primera fila en algunos archivos lleva nombres — se omite como en tu Tkinter
            continue
        row = []
        for col_idx, val in enumerate(record):
            var_name = headers_dict[col_idx]
            labels_for_var = value_labels_dict.get(var_name)
            if labels_for_var is not None and val in labels_for_var:
                row.append(labels_for_var[val])
            else:
                row.append(decode_bytes(val))
        rows.append(row)

    return rows, final_headers, value_labels_dict


def to_excel_bytes(headers: List[str], rows: List[List[Any]], sheet_name: str) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name
    ws.append(headers)
    for r in rows:
        ws.append(r)
    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()

# ======================
# UI – Streamlit
# ======================

def main():
    st.set_page_config(
        page_title="SPSS → Excel",
        page_icon="📊",
        layout="centered",
        initial_sidebar_state="expanded",
    )

    # Idioma en estado
    if "lang" not in st.session_state:
        st.session_state.lang = "es"

    # Sidebar
    with st.sidebar:
        st.markdown("### 🌐 Language / Idioma")
        toggle = st.toggle(TEXTS[st.session_state.lang]["toggle_lang"], value=False, key="lang_toggle")
        if toggle:
            st.session_state.lang = "en" if st.session_state.lang == "es" else "es"
        texts = TEXTS[st.session_state.lang]

        st.markdown("---")
        st.markdown("#### ⚙️ Opciones / Options")
        show_preview = st.checkbox("👀 " + ("Vista previa" if st.session_state.lang == "es" else "Preview"), True)
        preview_rows = st.number_input("Rows", min_value=5, max_value=200, value=30, step=5)

        st.markdown("---")
        st.info(texts["tips"])  # consejo útil

    # Header
    st.markdown(f"## {texts['title']}")
    st.markdown(texts["subtitle"]) 

    uploaded = st.file_uploader(texts["uploader"], type=["sav", "zsav"], accept_multiple_files=False)

    if not uploaded:
        st.empty()
        st.caption(texts["no_file"])
        return

    # Procesamiento con estado
    with st.status(texts["load_status"], expanded=False) as status:
        try:
            data_bytes = uploaded.getvalue()
            rows, headers, _ = process_sav(data_bytes)
            status.update(label=texts["success_load"], state="complete")
        except Exception as e:
            status.update(label=f"{texts['error_load']}: {e}", state="error")
            st.stop()

    # Info del archivo
    cols = st.columns(3)
    with cols[0]:
        st.metric("Columns", len(headers))
    with cols[1]:
        st.metric("Rows", len(rows))
    with cols[2]:
        st.metric("Size", f"{uploaded.size/1024/1024:.2f} MB")

    st.markdown(texts["file_info"].format(name=uploaded.name, rows=len(rows), cols=len(headers)))

    # Vista previa
    if show_preview and rows:
        # Construimos un pequeño DataFrame sin depender de pandas (Streamlit puede mostrar listas de diccionarios)
        try:
            import pandas as pd
            df = pd.DataFrame(rows[: int(preview_rows)], columns=headers)
            st.dataframe(df, use_container_width=True)
        except Exception:
            # fallback simple
            st.write({h: r for h, r in zip(headers, rows[0])})

    # Botón de descarga
    with st.spinner(texts["saving"]):
        excel_bytes = to_excel_bytes(headers, rows, texts["sheetname"])

    st.success(texts["success_save"], icon="✅")
    st.download_button(
        label=texts["download"],
        data=excel_bytes,
        file_name=Path(uploaded.name).with_suffix('.xlsx').name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )


if __name__ == "__main__":
    main()
