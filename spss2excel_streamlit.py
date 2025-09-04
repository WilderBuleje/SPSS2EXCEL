import collections
import collections.abc
collections.Iterable = collections.abc.Iterable  # Compat. savReaderWriter en Python 3.10+

import streamlit as st
from savReaderWriter import SavReader
from io import BytesIO
from pathlib import Path
import tempfile
from typing import Dict, Any, List, Tuple
import xlsxwriter  # reemplazo de openpyxl

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
def process_sav(file_bytes: bytes) -> Tuple[List[List[Any]], List[str]]:
    """Procesa el SAV con SavReader, aplica etiquetas de variables/valores y retorna (rows, headers)."""
    # Escribir bytes a un archivo temporal
    with tempfile.NamedTemporaryFile(delete=False, suffix=".sav") as tmp:
        tmp.write(file_bytes)
        tmp_path = tmp.name

    with SavReader(tmp_path, returnHeader=True, rawMode=False) as reader:
        header = reader.header
        varLabels = reader.varLabels
        valueLabels = reader.valueLabels
        records = reader.all()

    headers_dict = {i: decode_bytes(h) for i, h in enumerate(header)}

    var_labels_dict: Dict[str, str] = {
        decode_bytes(k): decode_bytes(v) for k, v in varLabels.items()
    } if varLabels else {}

    value_labels_dict: Dict[str, Dict[Any, str]] = {}
    if valueLabels:
        for var_key, val_dict in valueLabels.items():
            var_key_str = decode_bytes(var_key)
            value_labels_dict[var_key_str] = {
                val_code: decode_bytes(val_label) for val_code, val_label in val_dict.items()
            }

    # Headers finales con etiqueta
    final_headers: List[str] = []
    for i in range(len(headers_dict)):
        var_name = headers_dict[i]
        label = var_labels_dict.get(var_name, "").strip()
        final_headers.append(f"{var_name} ({label})" if label else var_name)

    # Filas con etiquetas aplicadas
    rows: List[List[Any]] = []
    for idx, record in enumerate(records):
        if idx == 0:
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

    return rows, final_headers


def to_excel_bytes(headers: List[str], rows: List[List[Any]], sheet_name: str) -> bytes:
    """Convierte a Excel usando XlsxWriter (más rápido y ligero que openpyxl)."""
    output = BytesIO()
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    worksheet = workbook.add_worksheet(sheet_name)

    # Escribir encabezados
    for col, h in enumerate(headers):
        worksheet.write(0, col, h)

    # Escribir filas
    for row_idx, row in enumerate(rows, start=1):
        for col_idx, val in enumerate(row):
            worksheet.write(row_idx, col_idx, val)

    workbook.close()
    output.seek(0)
    return output.getvalue()

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

    if "lang" not in st.session_state:
        st.session_state.lang = "es"

    with st.sidebar:
        toggle = st.toggle(TEXTS[st.session_state.lang]["toggle_lang"], value=False, key="lang_toggle")
        if toggle:
            st.session_state.lang = "en" if st.session_state.lang == "es" else "es"
        texts = TEXTS[st.session_state.lang]

        st.info(texts["tips"])

    texts = TEXTS[st.session_state.lang]

    st.markdown(f"## {texts['title']}")
    st.markdown(texts["subtitle"])

    uploaded = st.file_uploader(texts["uploader"], type=["sav", "zsav"], accept_multiple_files=False)

    if not uploaded:
        st.caption(texts["no_file"])
        return

    with st.status(texts["load_status"], expanded=False) as status:
        try:
            data_bytes = uploaded.getvalue()
            rows, headers = process_sav(data_bytes)
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
