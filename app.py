import streamlit as st
import io
import re
import zipfile
import pandas as pd
import PyPDF2
from pathlib import Path
from openpyxl.styles import Border, Side, Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from extractor_movimientos import parsear_archivo, crear_excel, generar_sifere_txt, generar_sifere_retenciones_txt, generar_percepciones_arba_txt, CONCEPTOS_MAP

# --- Page Config ---
st.set_page_config(
    page_title="ADDISYC ETL",
    page_icon="📗",
    layout="centered"
)

# --- Styling ---
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Space+Mono:wght@400;700&family=Syne:wght@400;600;800&display=swap');

:root {
    --bg:        #0d0f14;
    --surface:   #141720;
    --border:    #252935;
    --accent:    #e8c84a;
    --accent2:   #4ae8a0;
    --text:      #e4e8f0;
    --muted:     #6b7280;
    --danger:    #f87171;
    --radius:    10px;
}

*, *::before, *::after { box-sizing: border-box; }

.stApp {
    background-color: var(--bg) !important;
    font-family: 'Syne', sans-serif;
    color: var(--text);
}

.block-container {
    padding-top: 2.5rem !important;
    padding-bottom: 3rem !important;
    max-width: 860px !important;
}

h1, h2, h3, h4, p, span, div, label {
    color: var(--text) !important;
}

/* Header */
.etl-logo {
    font-family: 'Space Mono', monospace;
    font-size: 0.7rem;
    letter-spacing: 0.35em;
    color: var(--accent) !important;
    text-transform: uppercase;
    text-align: center;
    margin-bottom: 0.8rem;
}
.etl-title {
    font-family: 'Syne', sans-serif !important;
    font-weight: 800;
    font-size: 3.4rem !important;
    line-height: 1.4;
    color: var(--text) !important;
    text-align: center;
    margin: 0 0 0.5rem !important;
}
.etl-title span { color: var(--accent) !important; }
.etl-subtitle {
    font-size: 0.85rem;
    color: var(--muted) !important;
    font-family: 'Space Mono', monospace;
    letter-spacing: 0.05em;
    text-align: center;
}
.divider {
    border: none;
    border-top: 1px solid var(--border);
    margin: 1.8rem 0;
}

/* Cards */
.card {
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: var(--radius);
    padding: 1.6rem 1.8rem;
    margin-bottom: 1.2rem;
    position: relative;
    overflow: hidden;
}
.card::before {
    content: '';
    position: absolute;
    top: 0; left: 0; right: 0;
    height: 2px;
    background: linear-gradient(90deg, var(--accent), transparent);
}
.card-label {
    font-family: 'Space Mono', monospace;
    font-size: 0.75rem;
    font-weight: 700;
    letter-spacing: 0.2em;
    color: var(--accent) !important;
    text-transform: uppercase;
    margin-bottom: 1rem;
}

/* File uploader */
[data-testid="stFileUploader"] > div,
[data-testid="stFileUploader"] > div > div,
[data-testid="stFileUploader"] section,
[data-testid="stFileUploader"] section > div,
[data-testid="stFileUploadDropzone"],
[data-testid="stFileDropzoneInstructions"],
.stFileUploader > div,
.stFileUploader section {
    background: #1a1d24 !important;
    background-color: #1a1d24 !important;
    border: 1.5px dashed var(--border) !important;
    border-radius: var(--radius) !important;
    transition: border-color 0.2s ease;
}
[data-testid="stFileUploader"] > div:hover,
[data-testid="stFileUploadDropzone"]:hover,
.stFileUploader > div:hover {
    border-color: var(--accent) !important;
}
.stFileUploader label, [data-testid="stFileUploader"] label {
    color: var(--muted) !important;
}
[data-testid="stFileUploader"] small,
[data-testid="stFileUploader"] span,
[data-testid="stFileDropzoneInstructions"] span,
[data-testid="stFileDropzoneInstructions"] small,
[data-testid="stFileDropzoneInstructions"] div {
    color: var(--muted) !important;
}
.stFileUploader button, [data-testid="stFileUploader"] button {
    background: var(--surface) !important;
    color: var(--accent) !important;
    border: 1px solid var(--border) !important;
    border-radius: 6px !important;
    font-family: 'Space Mono', monospace !important;
    font-size: 0.75rem !important;
}
.stFileUploader button:hover, [data-testid="stFileUploader"] button:hover {
    border-color: var(--accent) !important;
}

/* Checkbox */
.stCheckbox label span { color: var(--text) !important; }
[data-testid="stCheckbox"] > label > div {
    border-color: var(--border) !important;
}

/* Main action button */
.stButton > button {
    width: 100% !important;
    background: var(--accent) !important;
    color: #0a0c10 !important;
    border: none !important;
    border-radius: var(--radius) !important;
    font-family: 'Syne', sans-serif !important;
    font-weight: 800 !important;
    font-size: 1rem !important;
    letter-spacing: 0.08em;
    height: 3.2em !important;
    margin-top: 0.5rem;
    transition: all 0.18s ease !important;
    box-shadow: 0 0 20px rgba(232,200,74,0.15);
    text-shadow: none !important;
    -webkit-text-fill-color: #0a0c10 !important;
}
.stButton > button:hover {
    background: #f5d84e !important;
    box-shadow: 0 0 30px rgba(232,200,74,0.3) !important;
    transform: translateY(-1px);
}
.stButton > button:active { transform: translateY(0); }

/* Download button */
[data-testid="stDownloadButton"] > button {
    background: transparent !important;
    color: var(--accent2) !important;
    border: 1.5px solid var(--accent2) !important;
    border-radius: var(--radius) !important;
    font-family: 'Space Mono', monospace !important;
    font-size: 0.8rem !important;
    letter-spacing: 0.06em;
    width: 100% !important;
    height: 3em !important;
    margin-top: 0.8rem;
    transition: all 0.18s ease !important;
}
[data-testid="stDownloadButton"] > button:hover {
    background: rgba(74,232,160,0.08) !important;
    box-shadow: 0 0 20px rgba(74,232,160,0.2) !important;
}

/* Alerts */
[data-testid="stAlert"] {
    border-radius: var(--radius) !important;
}
.stSuccess > div {
    background: rgba(74,232,160,0.07) !important;
    border: 1px solid rgba(74,232,160,0.25) !important;
}
.stSuccess p, .stSuccess span, .stSuccess strong { color: var(--accent2) !important; }

.stError > div {
    background: rgba(248,113,113,0.07) !important;
    border: 1px solid rgba(248,113,113,0.3) !important;
}
.stError p, .stError span { color: var(--danger) !important; }

.stWarning > div {
    background: rgba(232,200,74,0.07) !important;
    border: 1px solid rgba(232,200,74,0.25) !important;
}
.stWarning p, .stWarning span { color: var(--accent) !important; }

.stInfo > div {
    background: rgba(99,122,255,0.07) !important;
    border: 1px solid rgba(99,122,255,0.3) !important;
}
.stInfo p, .stInfo span, .stInfo strong { color: #a5b4fc !important; }

/* Spinner */
.stSpinner > div { border-top-color: var(--accent) !important; }

/* Stats row */
.stats-row {
    display: flex;
    gap: 0.8rem;
    margin-top: 1rem;
}
.stat-chip {
    flex: 1;
    background: #0a0c10;
    border: 1px solid var(--border);
    border-radius: 8px;
    padding: 0.7rem 0.5rem;
    text-align: center;
}
.stat-chip .stat-val {
    font-family: 'Space Mono', monospace;
    font-size: 1.3rem;
    font-weight: 700;
    color: var(--accent) !important;
    display: block;
}
.stat-chip .stat-lbl {
    font-size: 0.65rem;
    letter-spacing: 0.1em;
    color: var(--muted) !important;
    text-transform: uppercase;
    display: block;
    margin-top: 0.2rem;
}

/* Scrollbar */
::-webkit-scrollbar { width: 5px; }
::-webkit-scrollbar-track { background: var(--bg); }
::-webkit-scrollbar-thumb { background: var(--border); border-radius: 99px; }

/* Navbar / Header bar */
header[data-testid="stHeader"],
.stAppHeader,
header.stAppHeader {
    background: #1a1d24 !important;
    background-color: #1a1d24 !important;
    border-bottom: 1px solid var(--border) !important;
}
header[data-testid="stHeader"] *,
.stAppHeader * {
    color: var(--accent) !important;
}

/* Selectbox */
[data-testid="stSelectbox"] > div > div,
.stSelectbox > div > div {
    background: #1a1d24 !important;
    background-color: #1a1d24 !important;
    border: 1.5px solid var(--border) !important;
    border-radius: var(--radius) !important;
    color: var(--accent) !important;
}
[data-testid="stSelectbox"] > div > div:hover,
.stSelectbox > div > div:hover {
    border-color: var(--accent) !important;
}
[data-testid="stSelectbox"] span,
[data-testid="stSelectbox"] div[data-baseweb="select"] span,
.stSelectbox span {
    color: var(--accent) !important;
}
[data-testid="stSelectbox"] svg {
    fill: var(--accent) !important;
}
[data-testid="stSelectbox"] label {
    color: var(--muted) !important;
}
/* Selectbox dropdown menu */
[data-baseweb="popover"],
[data-baseweb="popover"] > div,
[data-baseweb="menu"],
ul[role="listbox"],
div[data-baseweb="popover"] div,
div[data-baseweb="popover"] ul {
    background: #1a1d24 !important;
    background-color: #1a1d24 !important;
    border-color: var(--border) !important;
}
[data-baseweb="popover"] {
    border: 1px solid var(--border) !important;
    border-radius: var(--radius) !important;
}
ul[role="listbox"] li,
[data-baseweb="menu"] li,
[data-baseweb="popover"] li,
li[role="option"] {
    background: #1a1d24 !important;
    background-color: #1a1d24 !important;
    color: var(--text) !important;
}
ul[role="listbox"] li:hover,
[data-baseweb="menu"] li:hover,
[data-baseweb="popover"] li:hover,
li[role="option"]:hover,
ul[role="listbox"] li[aria-selected="true"],
[data-baseweb="menu"] li[aria-selected="true"],
li[role="option"][aria-selected="true"] {
    background: var(--surface) !important;
    background-color: var(--surface) !important;
    color: var(--accent) !important;
}

/* Radio buttons */
[data-testid="stRadio"] > div > div > label > div {
    background: #1a1d24 !important;
}

/* Footer */
.etl-footer {
    text-align: center;
    padding-top: 2rem;
    font-family: 'Space Mono', monospace;
    font-size: 0.62rem;
    color: var(--muted) !important;
    letter-spacing: 0.15em;
}
</style>
""", unsafe_allow_html=True)


# ─── Header ────────────────────────────────────────────────────────────────────
st.markdown("""
<div>
    <h1 class="etl-title">Transformación<span> Datos Mendez</span></h1>
    <p class="etl-subtitle">TXT  →  XLSX</p>
</div>
<hr class="divider">
""", unsafe_allow_html=True)


# ─── Selector de herramienta ────────────────────────────────────────────────────────────
TOOL_MOVIMIENTOS = "Extracción de Movimientos (.txt)"
TOOL_PORTAL_IVA = "Movimientos Portal IVA limpio (.zip)"
TOOL_SIFERE = "Archivos SIFERE (.txt)"
TOOL_LIQUIDACIONES = "Liquidaciones Tarjeta (.pdf)"
TOOL_DEDUCCIONES = "Limpieza Excel Deducciones IVA/Ganancias"
TOOL_ARBA = "Archivo Agente de Percepciones ARBA (.txt)"
TOOL_CRUCE_CONCEPTO = "Cruce (TXT + Excel Sistema)"

herramienta = st.selectbox(
    "Seleccioná la herramienta:",
    options=[TOOL_MOVIMIENTOS, TOOL_PORTAL_IVA, TOOL_SIFERE, TOOL_ARBA, TOOL_LIQUIDACIONES, TOOL_DEDUCCIONES, TOOL_CRUCE_CONCEPTO],
    index=0,
)

if herramienta in (TOOL_MOVIMIENTOS, TOOL_PORTAL_IVA):
    with st.expander("📋 Códigos de comprobantes ARCA"):
        st.markdown("""
| Código | Tipo | Código | Tipo | Código | Tipo |
|--------|------|--------|------|--------|------|
| 1 | FC A | 2 | ND A | 3 | NC A |
| 6 | FC B | 7 | ND B | 8 | NC B |
| 11 | FC C | 12 | ND C | 13 | NC C |
| 51 | FC M | 52 | ND M | 53 | NC M |
| 19 | FC | 20 | ND | 21 | NC |
| 22 | FC | 37 | ND | 38 | NC |
| 195 | FC T | 196 | ND T | 197 | NC T |
| 201 | FC A | 202 | ND A | 203 | NC A |
| 206 | FC B | 207 | ND B | 208 | NC B |
| 211 | FC C | 212 | ND C | 213 | NC C |
| 81 | TF A | 45 | ND A | 48 | NC A |
| 82 | TF B | 46 | ND B | 43 | NC B |
| 111 | TF C | 47 | ND C | 44 | NC C |
| 118 | TF M | 90 | NC | 83 | TK |
| 109 | TK C | 110 | TK | 112 | TK A |
| 113 | TK B | 114 | TK C | 115 | TK A |
| 116 | TK B | 117 | TK C | 119 | TK M |
| 120 | TK M | 4 | RC A → FC A | 9 | RC B → FC B |
| 15 | RC C → FC C | | | | |

**FC** = Factura · **NC** = Nota de Crédito · **ND** = Nota de Débito · **TF** = Tique Factura · **TK** = Tique · **RC** = Recibo (se trata como FC)
        """)

if herramienta == TOOL_MOVIMIENTOS:
        # ─── Card 01: Archivo ──────────────────────────────────────────────────────────
        st.markdown('<div class="card"><div class="card-label">01 · Archivo fuente</div>', unsafe_allow_html=True)
        uploaded_file = st.file_uploader(
            "Arrastrá tu archivo o hacé click para seleccionarlo",
            type=["txt", "prn"],
            label_visibility="visible"
        )
        st.markdown('</div>', unsafe_allow_html=True)


        # ─── Card 02: Opciones ─────────────────────────────────────────────────────────
        st.markdown('<div class="card"><div class="card-label">02 · Opciones de exportación</div>', unsafe_allow_html=True)
        OPT_SOLO = "Solo Movimientos"
        OPT_AUXILIAR = "Exportar con columna Auxiliar"
        OPT_RESUMENES = "Incluir hojas de resumen"
        OPT_ARCA = "Cruce de comprobantes con ARCA"

        modo_export = st.radio(
            "Seleccioná el modo de exportación:",
            options=[OPT_SOLO, OPT_AUXILIAR, OPT_RESUMENES, OPT_ARCA],
            index=0,
            help="Solo se puede elegir una opción a la vez."
        )
        con_auxiliar = modo_export == OPT_AUXILIAR
        con_resumenes = modo_export == OPT_RESUMENES
        cruce_arca = modo_export == OPT_ARCA
        st.markdown('</div>', unsafe_allow_html=True)

        # ─── Card 02b: Archivo ARCA (condicional) ──────────────────────────────────────
        df_arca = None
        if cruce_arca:
            st.markdown('<div class="card"><div class="card-label">02b · Archivo ARCA (.zip)</div>', unsafe_allow_html=True)
            uploaded_arca = st.file_uploader(
                "Subí el .zip descargado de ARCA con los comprobantes",
                type=["zip"],
                label_visibility="visible",
                key="arca_zip"
            )
            if uploaded_arca:
                try:
                    with zipfile.ZipFile(io.BytesIO(uploaded_arca.getvalue())) as zf:
                        all_files = [f for f in zf.namelist() if not f.endswith('/')]
                        if all_files:
                            target_file = all_files[0]
                            with zf.open(target_file) as data_file:
                                raw = data_file.read()
                            csv_text = raw.decode('latin-1')
                            sep = ';' if csv_text.count(';') > csv_text.count(',') else ','
                            df_arca = pd.read_csv(
                                io.StringIO(csv_text), sep=sep, on_bad_lines='skip'
                            )
                            # Mapear códigos de comprobante ARCA a tipos del sistema (con letra)
                            ARCA_TIPO_MAP = {
                                # Facturas
                                1: 'FC A', 6: 'FC B', 11: 'FC C', 51: 'FC M',
                                19: 'FC', 22: 'FC', 195: 'FC T',
                                201: 'FC A', 206: 'FC B', 211: 'FC C',
                                # Recibos (se tratan como FC)
                                4: 'FC A', 9: 'FC B', 15: 'FC C',
                                # Notas de Débito
                                2: 'ND A', 7: 'ND B', 12: 'ND C', 52: 'ND M',
                                20: 'ND', 37: 'ND', 196: 'ND T',
                                45: 'ND A', 46: 'ND B', 47: 'ND C',
                                202: 'ND A', 207: 'ND B', 212: 'ND C',
                                # Notas de Crédito
                                3: 'NC A', 8: 'NC B', 13: 'NC C', 53: 'NC M',
                                21: 'NC', 38: 'NC', 90: 'NC', 197: 'NC T',
                                43: 'NC B', 44: 'NC C', 48: 'NC A',
                                203: 'NC A', 208: 'NC B', 213: 'NC C',
                                # Tique Factura
                                81: 'TF A', 82: 'TF B', 111: 'TF C', 118: 'TF M',
                                # Tique
                                83: 'TK', 109: 'TK C', 110: 'TK',
                                112: 'TK A', 113: 'TK B', 114: 'TK C',
                                115: 'TK A', 116: 'TK B', 117: 'TK C',
                                119: 'TK M', 120: 'TK M',
                            }
                            col_tipo = 'Tipo de Comprobante'
                            if col_tipo in df_arca.columns:
                                df_arca[col_tipo] = pd.to_numeric(df_arca[col_tipo], errors='coerce').astype('Int64')
                                df_arca[col_tipo] = df_arca[col_tipo].map(ARCA_TIPO_MAP).fillna(df_arca[col_tipo].astype(str))

                            # ── Limpieza de columnas ARCA ──────────────────────────────
                            # Renombrar columnas (usa partial match para encodings rotos)
                            RENAME_RULES = [
                                (['fecha', 'emisi'], 'Fecha'),
                                (['tipo', 'comprobante'], 'Comprobante'),
                                (['punto', 'venta'], 'PV'),
                                (['mero', 'comprobante'], 'Nro.'),
                                (['tipo', 'doc', 'vendedor'], 'Tipo Doc.'),
                                (['nro', 'doc', 'vendedor'], 'CUIT'),
                                (['denominaci', 'vendedor'], 'Razon Social'),
                                (['importe', 'total'], 'Total'),
                        (['moneda', 'original'], 'Moneda'),
                        (['tipo', 'cambio'], 'Tipo Cambio'),
                                (['importe', 'no', 'gravado'], 'No Gravado'),
                                (['importe', 'exento'], 'Exento'),
                                (['pagos', 'cta', 'otros'], 'Otras Perc.'),
                                (['percepciones', 'ingresos', 'brutos'], 'Perc IIBB'),
                                (['impuestos', 'municipales'], 'Impuestos Munic.'),
                                (['percepciones', 'pagos', 'cuenta', 'iva'], 'Perc. IVA'),
                                (['impuestos', 'internos'], 'Imp. Int.'),
                                (['importe', 'otros', 'tributos'], 'Otros. Trib.'),
                                (['neto', 'gravado', 'iva', '0'], 'IVA 0%'),
                                (['neto', 'gravado', 'iva', '21'], 'Gravado IVA 21'),
                                (['importe', 'iva', '21'], 'IVA 21'),
                                (['neto', 'gravado', 'iva', '27'], 'Gravado IVA 27'),
                                (['importe', 'iva', '27'], 'IVA 27'),
                                (['neto', 'gravado', 'iva', '10'], 'Gravado IVA 10,5'),
                                (['importe', 'iva', '10'], 'IVA 10,5'),
                                (['neto', 'gravado', 'iva', '2'], 'Gravado IVA 2,5'),
                                (['importe', 'iva', '2'], 'IVA 2,5'),
                                (['neto', 'gravado', 'iva', '5%'], 'Gravado IVA 5'),
                                (['importe', 'iva', '5%'], 'IVA 5'),
                            ]
                            rename_map = {}
                            for keywords, new_name in RENAME_RULES:
                                for c in df_arca.columns:
                                    cl = c.strip().lower()
                                    if all(k in cl for k in keywords) and c not in rename_map:
                                        rename_map[c] = new_name
                                        break
                            df_arca = df_arca.rename(columns=rename_map)

                            # Convertir fecha de aaaa-mm-dd a dd/mm/aaaa
                            if 'Fecha' in df_arca.columns:
                                df_arca['Fecha'] = df_arca['Fecha'].astype(str).apply(
                                    lambda x: '/'.join(x.split('-')[::-1]) if '-' in x else x
                                )

                            # Eliminar columnas no deseadas
                            DROP_KEYWORDS = [
                                ['dito', 'fiscal', 'computable'],
                                ['total', 'neto', 'gravado'],
                                ['total', 'iva'],
                                ['tipo', 'doc'],
                            ]
                            cols_to_drop = []
                            for keywords in DROP_KEYWORDS:
                                for c in df_arca.columns:
                                    cl = c.strip().lower()
                                    if all(k in cl for k in keywords):
                                        cols_to_drop.append(c)
                                        break
                            df_arca = df_arca.drop(columns=[c for c in cols_to_drop if c in df_arca.columns], errors='ignore')

                            # Mover Total al final
                            if 'Total' in df_arca.columns:
                                total_data = df_arca.pop('Total')
                                df_arca['Total'] = total_data

                            # Columna Auxiliar: Tipo + PV + Nro Comprobante + Nro Doc Vendedor
                            def find_col(df, keywords):
                                """Busca columna que contenga todas las keywords (case-insensitive)."""
                                for c in df.columns:
                                    cl = c.strip().lower()
                                    if all(k in cl for k in keywords):
                                        return c
                                return None

                            # Crear columna Auxiliar con nombres ya renombrados
                            aux_cols = ['Comprobante', 'PV', 'Nro.', 'CUIT']
                            if all(c in df_arca.columns for c in aux_cols):
                                df_arca['Auxiliar'] = (
                                    df_arca['Comprobante'].astype(str) +
                                    df_arca['PV'].astype(str) +
                                    df_arca['Nro.'].astype(str) +
                                    df_arca['CUIT'].astype(str)
                                )
                                # Mover Auxiliar justo antes de Total
                                cols = list(df_arca.columns)
                                cols.remove('Auxiliar')
                                total_pos = cols.index('Total') if 'Total' in cols else len(cols)
                                cols.insert(total_pos, 'Auxiliar')
                                df_arca = df_arca[cols]

                            # Columnas monetarias: desde 'No Gravado' en adelante (excluyendo Auxiliar)
                            all_cols = list(df_arca.columns)
                            ng_idx = all_cols.index('No Gravado') if 'No Gravado' in all_cols else None
                            if ng_idx is not None:
                                money_cols = [c for c in all_cols[ng_idx:] if c != 'Auxiliar']
                                for c in money_cols:
                                    # Convertir formato argentino: 1.234,56 -> 1234.56
                                    df_arca[c] = df_arca[c].astype(str).str.replace('.', '', regex=False).str.replace(',', '.', regex=False)
                                    df_arca[c] = pd.to_numeric(df_arca[c], errors='coerce').fillna(0)
                                # Eliminar columnas monetarias que son todo cero
                                empty_money = [c for c in money_cols if (df_arca[c] == 0).all()]
                                df_arca = df_arca.drop(columns=empty_money)

                            st.success(f"**{target_file}** · {len(df_arca)} comprobantes leídos de ARCA")
                        else:
                            st.error("El .zip está vacío")
                except Exception as e:
                    st.error(f"Error al leer el .zip: {str(e)}")
            else:
                st.info("Subí el archivo .zip de ARCA para continuar")
            st.markdown('</div>', unsafe_allow_html=True)


        if uploaded_file is not None:
            filename = Path(uploaded_file.name).stem
            st.success(f"**{uploaded_file.name}** listo para procesar")

            st.markdown('<div class="card"><div class="card-label">03 · Procesar</div>', unsafe_allow_html=True)

            if st.button("⬡  Procesar Archivo"):
                try:
                    with st.spinner("Analizando información..."):
                        content = uploaded_file.getvalue().decode("latin-1")
                        transacciones, meta = parsear_archivo(content=content)

                    if not transacciones:
                        st.error("No se encontraron transacciones. Verificá el formato del archivo.")
                    else:
                        with st.spinner("Generando Excel..."):
                            output = io.BytesIO()
                            crear_excel(transacciones, meta, output,
                                        con_resumenes=con_resumenes,
                                        con_auxiliar=con_auxiliar,
                                        cruce_arca=cruce_arca,
                                        df_arca=df_arca)
                            output.seek(0)

                        st.success("✓  Proceso completado con éxito")

                        # Stats chips
                        from collections import Counter
                        tipos = Counter(t['Tipo'] for t in transacciones)
                        st.markdown(f"""
                        <div class="stats-row">
                            <div class="stat-chip">
                                <span class="stat-val">{len(transacciones)}</span>
                                <span class="stat-lbl">Total</span>
                            </div>
                            <div class="stat-chip">
                                <span class="stat-val">{tipos.get('FC', 0)}</span>
                                <span class="stat-lbl">Facturas</span>
                            </div>
                            <div class="stat-chip">
                                <span class="stat-val">{tipos.get('NC', 0)}</span>
                                <span class="stat-lbl">Notas Cred.</span>
                            </div>
                            <div class="stat-chip">
                                <span class="stat-val">{tipos.get('ND', 0) + tipos.get('TF', 0) + tipos.get('TK', 0)}</span>
                                <span class="stat-lbl">Otros</span>
                            </div>
                        </div>
                        """, unsafe_allow_html=True)

                        st.info(
                            f"**{meta.get('tipo_reporte', 'N/A')}** · "
                            f"{meta.get('razon_social', 'Contribuyente')} · "
                            f"{meta.get('periodo', '')}"
                        )

                        excel_filename = "Cruce Compras.xlsx" if cruce_arca else f"{filename}_procesado.xlsx"
                        st.download_button(
                            label="↓  Descargar Excel",
                            data=output,
                            file_name=excel_filename,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True,
                        )

                except Exception as e:
                    st.error(f"Error al procesar el archivo: {str(e)}")
                    st.exception(e)

            st.markdown('</div>', unsafe_allow_html=True)

        else:
            st.markdown("""
            <div style="
                text-align: center;
                padding: 2rem 1rem;
                font-family: 'Space Mono', monospace;
                font-size: 0.72rem;
                color: #6b7280;
                letter-spacing: 0.12em;
            ">
                ESPERANDO ARCHIVO · PASO 01
            </div>
            """, unsafe_allow_html=True)


elif herramienta == TOOL_PORTAL_IVA:
    # ───────────────────────────────────────────────────────────────────────────────
    # HERRAMIENTA: Portal IVA limpio
    # ───────────────────────────────────────────────────────────────────────────────
    st.markdown('<div class="card"><div class="card-label">01 · Archivo ARCA (.zip)</div>', unsafe_allow_html=True)
    uploaded_zip_iva = st.file_uploader(
        "Subí el .zip descargado del Portal IVA de ARCA",
        type=["zip"],
        label_visibility="visible",
        key="portal_iva_zip"
    )
    st.markdown('</div>', unsafe_allow_html=True)

    if uploaded_zip_iva:
        st.success(f"**{uploaded_zip_iva.name}** listo para procesar")

        st.markdown('<div class="card"><div class="card-label">02 · Datos del contribuyente</div>', unsafe_allow_html=True)
        nombre_contribuyente = st.text_input("Nombre / Razón Social del contribuyente", value="", key="nombre_portal_iva")
        st.markdown('</div>', unsafe_allow_html=True)

        st.markdown('<div class="card"><div class="card-label">03 · Procesar</div>', unsafe_allow_html=True)

        if st.button("⬡  Procesar ZIP"):
            if not nombre_contribuyente.strip():
                st.error("Ingresá el nombre del contribuyente para continuar.")
            else:
              try:
                with st.spinner("Leyendo archivo ARCA..."):
                    with zipfile.ZipFile(io.BytesIO(uploaded_zip_iva.getvalue())) as zf:
                        all_files = [f for f in zf.namelist() if not f.endswith('/')]
                        if not all_files:
                            st.error("El .zip está vacío")
                            st.stop()
                        target_file = all_files[0]
                        with zf.open(target_file) as data_file:
                            raw = data_file.read()

                    csv_text = raw.decode('latin-1')
                    sep = ';' if csv_text.count(';') > csv_text.count(',') else ','
                    df_iva = pd.read_csv(io.StringIO(csv_text), sep=sep, on_bad_lines='skip')

                    # Detectar tipo (Compras/Ventas), CUIT y periodo del nombre del zip
                    import re as _re
                    zip_name_raw = uploaded_zip_iva.name.upper()
                    es_ventas = 'VENTA' in zip_name_raw
                    es_compras = 'COMPRA' in zip_name_raw
                    tipo_portal = 'VENTAS' if es_ventas else ('COMPRAS' if es_compras else 'PORTAL IVA')

                    # Buscar CUIT (11 dígitos) y periodo (YYYYMM o YYYY-MM)
                    cuit_match = _re.search(r'(\d{11})', zip_name_raw)
                    cuit_portal = cuit_match.group(1) if cuit_match else ''
                    periodo_match = _re.search(r'(\d{4})(\d{2})(?!\d)', zip_name_raw)
                    if periodo_match:
                        meses = ['','Enero','Febrero','Marzo','Abril','Mayo','Junio',
                                 'Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre']
                        m_num = int(periodo_match.group(2))
                        periodo_portal = f"{meses[m_num]} {periodo_match.group(1)}" if 1 <= m_num <= 12 else ''
                    else:
                        periodo_portal = ''

                    # Mapear códigos de comprobante
                    ARCA_TIPO_MAP = {
                        1: 'FC A', 6: 'FC B', 11: 'FC C', 51: 'FC M',
                        19: 'FC', 22: 'FC', 195: 'FC T',
                        201: 'FC A', 206: 'FC B', 211: 'FC C',
                        # Recibos (se tratan como FC)
                        4: 'FC A', 9: 'FC B', 15: 'FC C',
                        2: 'ND A', 7: 'ND B', 12: 'ND C', 52: 'ND M',
                        20: 'ND', 37: 'ND', 196: 'ND T',
                        45: 'ND A', 46: 'ND B', 47: 'ND C',
                        202: 'ND A', 207: 'ND B', 212: 'ND C',
                        3: 'NC A', 8: 'NC B', 13: 'NC C', 53: 'NC M',
                        21: 'NC', 38: 'NC', 90: 'NC', 197: 'NC T',
                        43: 'NC B', 44: 'NC C', 48: 'NC A',
                        203: 'NC A', 208: 'NC B', 213: 'NC C',
                        81: 'TF A', 82: 'TF B', 111: 'TF C', 118: 'TF M',
                        83: 'TK', 109: 'TK C', 110: 'TK',
                        112: 'TK A', 113: 'TK B', 114: 'TK C',
                        115: 'TK A', 116: 'TK B', 117: 'TK C',
                        119: 'TK M', 120: 'TK M',
                    }

                    def find_col_iva(df, keywords):
                        for c in df.columns:
                            cl = c.strip().lower()
                            if all(k in cl for k in keywords):
                                return c
                        return None

                    col_tipo_iva = find_col_iva(df_iva, ['tipo', 'comprobante'])
                    if col_tipo_iva:
                        df_iva[col_tipo_iva] = pd.to_numeric(df_iva[col_tipo_iva], errors='coerce').astype('Int64')
                        df_iva[col_tipo_iva] = df_iva[col_tipo_iva].map(ARCA_TIPO_MAP).fillna(df_iva[col_tipo_iva].astype(str))

                    # Renombrar columnas (funciona para compras y ventas)
                    RENAME_RULES = [
                        (['fecha', 'emisi'], 'Fecha'),
                        (['tipo', 'comprobante'], 'Comprobante'),
                        (['punto', 'venta'], 'PV'),
                        (['mero', 'comprobante', 'hasta'], 'Nro. Hasta'),
                        (['mero', 'comprobante'], 'Nro.'),
                        (['tipo', 'doc'], 'Tipo Doc.'),
                        (['nro', 'doc', 'vendedor'], 'CUIT'),
                        (['nro', 'doc', 'comprador'], 'CUIT'),
                        (['denominaci', 'vendedor'], 'Razon Social'),
                        (['denominaci', 'comprador'], 'Razon Social'),
                        (['fecha', 'vencimiento'], 'Fecha Vto. Pago'),
                        (['importe', 'total'], 'Total'),
                        (['moneda', 'original'], 'Moneda'),
                        (['tipo', 'cambio'], 'Tipo Cambio'),
                        (['importe', 'no', 'gravado'], 'No Gravado'),
                        (['importe', 'exento'], 'Exento'),
                        (['pagos', 'cta', 'otros'], 'Otras Perc.'),
                        (['percepciones', 'ingresos', 'brutos'], 'Perc IIBB'),
                        (['impuestos', 'municipales'], 'Impuestos Munic.'),
                        (['percepciones', 'pagos', 'cuenta', 'iva'], 'Perc. IVA'),
                        (['percepci', 'no', 'categorizados'], 'Perc. No Cat.'),
                        (['impuestos', 'internos'], 'Imp. Int.'),
                        (['importe', 'otros', 'tributos'], 'Otros. Trib.'),
                        (['neto', 'gravado', 'iva', '0'], 'IVA 0%'),
                        (['neto', 'gravado', 'iva', '21'], 'Gravado IVA 21'),
                        (['importe', 'iva', '21'], 'IVA 21'),
                        (['neto', 'gravado', 'iva', '27'], 'Gravado IVA 27'),
                        (['importe', 'iva', '27'], 'IVA 27'),
                        (['neto', 'gravado', 'iva', '10'], 'Gravado IVA 10,5'),
                        (['importe', 'iva', '10'], 'IVA 10,5'),
                        (['neto', 'gravado', 'iva', '2'], 'Gravado IVA 2,5'),
                        (['importe', 'iva', '2'], 'IVA 2,5'),
                        (['neto', 'gravado', 'iva', '5%'], 'Gravado IVA 5'),
                        (['importe', 'iva', '5%'], 'IVA 5'),
                    ]
                    rename_map = {}
                    for keywords, new_name in RENAME_RULES:
                        for c in df_iva.columns:
                            cl = c.strip().lower()
                            if all(k in cl for k in keywords) and c not in rename_map:
                                rename_map[c] = new_name
                                break
                    df_iva = df_iva.rename(columns=rename_map)

                    # Convertir fecha de aaaa-mm-dd a dd/mm/aaaa
                    if 'Fecha' in df_iva.columns:
                        df_iva['Fecha'] = df_iva['Fecha'].astype(str).apply(
                            lambda x: '/'.join(x.split('-')[::-1]) if '-' in x else x
                        )

                    # Eliminar columnas no deseadas
                    DROP_KW = [
                        ['dito', 'fiscal', 'computable'],
                        ['total', 'neto', 'gravado'],
                        ['total', 'iva'],
                        ['tipo', 'doc'],
                        ['nro.', 'hasta'],
                        ['fecha', 'vto'],
                    ]
                    cols_to_drop = []
                    for kws in DROP_KW:
                        for c in df_iva.columns:
                            cl = c.strip().lower()
                            if all(k in cl for k in kws):
                                cols_to_drop.append(c)
                                break
                    df_iva = df_iva.drop(columns=[c for c in cols_to_drop if c in df_iva.columns], errors='ignore')

                    # Mover Total al final
                    if 'Total' in df_iva.columns:
                        total_data = df_iva.pop('Total')
                        df_iva['Total'] = total_data

                    # Columnas monetarias: convertir y limpiar
                    all_cols_iva = list(df_iva.columns)
                    non_money = {'Fecha', 'Comprobante', 'PV', 'Nro.', 'CUIT', 'Razon Social', 'Moneda', 'Tipo Cambio'}
                    money_cols_iva = [c for c in all_cols_iva if c not in non_money and c in df_iva.select_dtypes(include='object').columns]
                    for c in money_cols_iva:
                        df_iva[c] = df_iva[c].astype(str).str.replace('.', '', regex=False).str.replace(',', '.', regex=False)
                        df_iva[c] = pd.to_numeric(df_iva[c], errors='coerce').fillna(0)
                    # Rellenar NaN restantes en columnas numéricas
                    for c in all_cols_iva:
                        if c not in non_money and df_iva[c].dtype in ('float64', 'int64'):
                            df_iva[c] = df_iva[c].fillna(0)
                    # Eliminar columnas monetarias todo cero
                    empty_cols = [c for c in all_cols_iva if c not in non_money and c in df_iva.columns and df_iva[c].dtype in ('float64', 'int64') and (df_iva[c] == 0).all()]
                    df_iva = df_iva.drop(columns=empty_cols)

                with st.spinner("Generando Excel..."):
                    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
                    from openpyxl.utils import get_column_letter

                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df_iva.to_excel(writer, sheet_name=tipo_portal, index=False, startrow=5)
                        ws = writer.sheets[tipo_portal]
                        n_cols = len(df_iva.columns)

                        title_font = Font(bold=True, size=14, color='FFFFFF')
                        title_fill = PatternFill('solid', fgColor='2F5496')
                        header_font = Font(bold=True, size=10, color='FFFFFF')
                        header_fill = PatternFill('solid', fgColor='4472C4')
                        header_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
                        center_align = Alignment(horizontal='center', vertical='center')
                        thin_border = Border(
                            left=Side(style='thin'), right=Side(style='thin'),
                            top=Side(style='thin'), bottom=Side(style='thin')
                        )
                        zebra_fill = PatternFill('solid', fgColor='D6E4F0')
                        money_fmt = '$#,##0.00'

                        ws.merge_cells(f'A1:{get_column_letter(n_cols)}1')
                        ws['A1'] = nombre_contribuyente.strip().upper()
                        ws['A1'].font = title_font; ws['A1'].fill = title_fill
                        ws['A1'].alignment = center_align

                        ws.merge_cells(f'A2:{get_column_letter(n_cols)}2')
                        sub_parts = [p for p in [f'CUIT: {cuit_portal}' if cuit_portal else '', f'{len(df_iva)} comprobantes'] if p]
                        ws['A2'] = ' | '.join(sub_parts)
                        ws['A2'].font = Font(bold=True, size=11, color='2F5496')
                        ws['A2'].alignment = center_align

                        ws.merge_cells(f'A5:{get_column_letter(n_cols)}5')
                        ws['A5'] = f'{tipo_portal} {periodo_portal}'.strip()
                        ws['A5'].font = Font(bold=True, size=12, color='2F5496')
                        ws['A5'].alignment = center_align

                        col_list = list(df_iva.columns)
                        non_money_set = {'Fecha', 'Comprobante', 'PV', 'Nro.', 'CUIT', 'Razon Social', 'Auxiliar'}
                        money_indices = [i + 1 for i, c in enumerate(col_list) if c not in non_money_set and df_iva[c].dtype in ('float64', 'int64')]

                        for col_idx in range(1, n_cols + 1):
                            cell = ws.cell(row=6, column=col_idx)
                            cell.font = header_font; cell.fill = header_fill
                            cell.alignment = header_align; cell.border = thin_border

                        for row_idx in range(7, len(df_iva) + 7):
                            for col_idx in range(1, n_cols + 1):
                                cell = ws.cell(row=row_idx, column=col_idx)
                                cell.alignment = center_align
                                if col_idx in money_indices:
                                    cell.number_format = money_fmt
                            if (row_idx - 7) % 2 == 0:
                                for col_idx in range(1, n_cols + 1):
                                    ws.cell(row=row_idx, column=col_idx).fill = zebra_fill

                        # Fila TOTAL con fórmulas SUM
                        total_row = len(df_iva) + 7
                        col_list = list(df_iva.columns)
                        non_money_set2 = {'Fecha', 'Comprobante', 'PV', 'Nro.', 'CUIT', 'Razon Social', 'Moneda', 'Tipo Cambio'}
                        for col_idx in range(1, n_cols + 1):
                            cell = ws.cell(row=total_row, column=col_idx)
                            col_name = col_list[col_idx - 1] if col_idx - 1 < len(col_list) else ''
                            if col_name not in non_money_set2 and col_idx in money_indices:
                                letter = get_column_letter(col_idx)
                                cell.value = f'=SUM({letter}7:{letter}{total_row - 1})'
                                cell.number_format = money_fmt
                            elif col_idx == 1:
                                cell.value = 'TOTAL'
                            cell.font = Font(bold=True, size=10, color='FFFFFF')
                            cell.fill = PatternFill('solid', fgColor='2F5496')
                            cell.alignment = center_align

                        for col_idx in range(1, n_cols + 1):
                            max_len = max(
                                len(str(ws.cell(row=r, column=col_idx).value or ''))
                                for r in range(6, min(len(df_iva) + 7, 50))
                            )
                            letter = get_column_letter(col_idx)
                            ws.column_dimensions[letter].width = max(max_len + 3, 8)

                    output.seek(0)

                st.success("✓  Proceso completado con éxito")

                from collections import Counter
                tipos_iva = Counter(df_iva['Comprobante']) if 'Comprobante' in df_iva.columns else {}
                fc_count = sum(v for k, v in tipos_iva.items() if str(k).startswith('FC'))
                nc_count = sum(v for k, v in tipos_iva.items() if str(k).startswith('NC'))
                otros_count = len(df_iva) - fc_count - nc_count

                st.markdown(f"""
                <div class="stats-row">
                    <div class="stat-chip">
                        <span class="stat-val">{len(df_iva)}</span>
                        <span class="stat-lbl">Total</span>
                    </div>
                    <div class="stat-chip">
                        <span class="stat-val">{fc_count}</span>
                        <span class="stat-lbl">Facturas</span>
                    </div>
                    <div class="stat-chip">
                        <span class="stat-val">{nc_count}</span>
                        <span class="stat-lbl">Notas Cred.</span>
                    </div>
                    <div class="stat-chip">
                        <span class="stat-val">{otros_count}</span>
                        <span class="stat-lbl">Otros</span>
                    </div>
                </div>
                """, unsafe_allow_html=True)

                excel_name = 'Compras' if tipo_portal == 'COMPRAS' else ('Ventas' if tipo_portal == 'VENTAS' else Path(uploaded_zip_iva.name).stem)
                st.download_button(
                    label="↓  Descargar Excel",
                    data=output,
                    file_name=f"{excel_name}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

              except Exception as e:
                st.error(f"Error al procesar: {str(e)}")
                st.exception(e)

        st.markdown('</div>', unsafe_allow_html=True)
    else:
        st.markdown("""
        <div style="
            text-align: center;
            padding: 2rem 1rem;
            font-family: 'Space Mono', monospace;
            font-size: 0.72rem;
            color: #6b7280;
            letter-spacing: 0.12em;
        ">
            ESPERANDO ARCHIVO .ZIP · PASO 01
        </div>
        """, unsafe_allow_html=True)


elif herramienta == TOOL_SIFERE:
    # ───────────────────────────────────────────────────────────────────────────────
    # HERRAMIENTA: Archivos SIFERE (TXT)
    # ───────────────────────────────────────────────────────────────────────────────
    st.markdown('<div class="card"><div class="card-label">01 · Tipo de archivo SIFERE</div>', unsafe_allow_html=True)
    tipo_sifere = st.radio(
        "¿Qué tipo de archivo SIFERE querés generar?",
        options=["Percepciones", "Retenciones"],
        horizontal=True,
        key="sifere_tipo"
    )
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="card"><div class="card-label">02 · Archivo fuente para SIFERE</div>', unsafe_allow_html=True)
    uploaded_sifere = st.file_uploader(
        "Arrastrá tu archivo de movimientos o hacé click para seleccionarlo",
        type=["txt", "prn"],
        label_visibility="visible",
        key="sifere_file"
    )
    st.markdown('</div>', unsafe_allow_html=True)

    if uploaded_sifere:
        sifere_filename = Path(uploaded_sifere.name).stem
        st.success(f"**{uploaded_sifere.name}** listo para procesar")

        tipo_label = tipo_sifere.lower()  # "percepciones" o "retenciones"
        st.markdown(f'<div class="card"><div class="card-label">03 · Generar TXT SIFERE ({tipo_sifere})</div>', unsafe_allow_html=True)

        if st.button(f"⬡  Generar archivo SIFERE ({tipo_sifere})"):
            try:
                with st.spinner("Procesando..."):
                    raw_bytes = uploaded_sifere.getvalue()
                    content_str = raw_bytes.decode('latin-1', errors='replace')
                    movimientos, metadata = parsear_archivo(content=content_str)

                    if tipo_sifere == "Percepciones":
                        txt_sifere = generar_sifere_txt(movimientos, metadata)
                    else:
                        txt_sifere = generar_sifere_retenciones_txt(movimientos, metadata)

                st.success(f"✓  Archivo SIFERE ({tipo_sifere}) generado con éxito")

                # Stats
                st.markdown(f"""
                <div class="stats-row">
                    <div class="stat-chip">
                        <span class="stat-val">{len(movimientos)}</span>
                        <span class="stat-lbl">Movimientos</span>
                    </div>
                    <div class="stat-chip">
                        <span class="stat-val">{len(txt_sifere.splitlines())}</span>
                        <span class="stat-lbl">Líneas TXT</span>
                    </div>
                </div>
                """, unsafe_allow_html=True)

                st.download_button(
                    label=f"↓  Descargar TXT SIFERE ({tipo_sifere})",
                    data=txt_sifere.encode("latin-1", errors="replace"),
                    file_name=f"{sifere_filename}_sifere_{tipo_label}.txt",
                    mime="text/plain",
                    use_container_width=True,
                )

            except Exception as e:
                st.error(f"Error al procesar el archivo: {str(e)}")
                st.exception(e)

        st.markdown('</div>', unsafe_allow_html=True)

    else:
        st.markdown("""
        <div style="
            text-align: center;
            padding: 2rem 1rem;
            font-family: 'Space Mono', monospace;
            font-size: 0.72rem;
            color: #6b7280;
            letter-spacing: 0.12em;
        ">
            ESPERANDO ARCHIVO · PASO 01
        </div>
        """, unsafe_allow_html=True)


elif herramienta == TOOL_LIQUIDACIONES:
    # ───────────────────────────────────────────────────────────────────────────────
    # HERRAMIENTA: Liquidaciones Tarjeta (PDF)
    # ───────────────────────────────────────────────────────────────────────────────
    st.markdown('<div class="card"><div class="card-label">01 · Archivo PDF de Liquidaciones</div>', unsafe_allow_html=True)
    uploaded_liq = st.file_uploader(
        "Arrastrá tu PDF de liquidaciones o hacé click para seleccionarlo",
        type=["pdf"],
        label_visibility="visible",
        key="liquidaciones_pdf"
    )
    st.markdown('</div>', unsafe_allow_html=True)

    if uploaded_liq:
        liq_filename = Path(uploaded_liq.name).stem
        st.success(f"**{uploaded_liq.name}** listo para procesar")

        # ─── Card 02: Datos del contribuyente ──────────────────────────────────────
        st.markdown('<div class="card"><div class="card-label">02 · Datos del contribuyente</div>', unsafe_allow_html=True)
        col_a, col_b = st.columns(2)
        with col_a:
            nombre_contribuyente = st.text_input(
                "Nombre del contribuyente",
                value="",
                placeholder="Ej: Juan Pérez",
                key="liq_contribuyente"
            )
        with col_b:
            tipo_tarjeta = st.selectbox(
                "Tipo de tarjeta / Entidad",
                options=["Visa Crédito", "Visa Débito", "Mastercard Crédito", "Mastercard Débito",
                         "American Express Crédito", "American Express Débito",
                         "Maestro Crédito", "Maestro Débito",
                         "Cabal Crédito", "Cabal Débito", "Naranja",
                         "First Data", "Otra"],
                index=0,
                key="liq_tarjeta"
            )
        # Si selecciona "Otra", mostrar text input
        if tipo_tarjeta == "Otra":
            tipo_tarjeta_custom = st.text_input(
                "Especificá el tipo de tarjeta / entidad",
                value="",
                placeholder="Ej: Mercado Pago",
                key="liq_tarjeta_custom"
            )
            tipo_tarjeta_final = tipo_tarjeta_custom.strip() if tipo_tarjeta_custom.strip() else "Otra"
        else:
            tipo_tarjeta_final = tipo_tarjeta

        periodo_liq = st.text_input(
            "Periodo (MM/AAAA)",
            value="",
            placeholder="Ej: 04/2025",
            key="liq_periodo"
        )
        if periodo_liq and not re.match(r'^(0[1-9]|1[0-2])/\d{4}$', periodo_liq):
            st.error("El periodo debe tener formato MM/AAAA (ej: 04/2025)")
        st.markdown('</div>', unsafe_allow_html=True)

        # ─── Card 03: Procesar ─────────────────────────────────────────────────────
        st.markdown('<div class="card"><div class="card-label">03 · Generar Excel de Liquidaciones</div>', unsafe_allow_html=True)

        btn_procesar = st.button("⬡  Procesar Liquidaciones")

        # ─── Botón Procesar ────────────────────────────────────────────────────────
        if btn_procesar:
            if not nombre_contribuyente.strip():
                st.warning("Ingresá el nombre del contribuyente antes de procesar.")
            elif not periodo_liq.strip():
                st.warning("Ingresá el periodo (MM/AAAA) antes de procesar.")
            else:
                try:
                    # Parsear periodo para el encabezado (MM/AAAA -> AAMM)
                    periodo_parts = periodo_liq.strip().split("/")
                    if len(periodo_parts) == 2:
                        mes_liq = periodo_parts[0].zfill(2)
                        anio_liq = periodo_parts[1]
                        periodo_codigo = anio_liq[2:] + mes_liq  # AAMM
                    else:
                        mes_liq = "01"
                        anio_liq = "2025"
                        periodo_codigo = "2501"

                    with st.spinner("Leyendo PDF..."):
                        reader = PyPDF2.PdfReader(io.BytesIO(uploaded_liq.getvalue()))
                        texto = "".join(page.extract_text() + "\n" for page in reader.pages)
                        texto_lines = texto.splitlines()

                    with st.spinner("Extrayendo movimientos..."):
                        capturar = False
                        movimientos = []
                        movimiento = {}
                        # Extraer nombre del banco de la segunda línea del PDF
                        banco = texto_lines[1] if len(texto_lines) > 1 else "Banco desconocido"

                        for linea in texto_lines:
                            if "F.de Pago" in linea:
                                capturar = False
                                cbu_match = re.search(r"\d{1,3}(\.\d{3})*,\d+\-?\s+(\d+)", linea)
                                nro_cbu = cbu_match.group(1) if cbu_match else "No se encontró Número de Liquidación"
                                fecha_match = re.search(r"(\d{2}/\d{2}/\d{4})", linea.split("Nro. Liq:")[1]) if "Liq:" in linea else None
                                if cbu_match:
                                    movimiento["Liquidacion"] = cbu_match.group(2) + ".00"
                                fecha_liq = fecha_match.group(1) if fecha_match else "No se encontró fecha"
                                movimiento["Fecha"] = fecha_liq
                                if nro_cbu != "No se encontró Número de Liquidación":
                                    movimiento["Liquidacion"] = round(float(movimiento["Liquidacion"]))
                                    movimientos.append(movimiento.copy())
                                    movimiento = {}

                            if "VENTAS" in linea or "QR" in linea or "AJUSTE" in linea or "ACREDITACIONES PAGO QRD" in linea:
                                capturar = True

                            if capturar:
                                partes = linea.split("$")
                                if len(partes) > 1:
                                    valor = partes[1].strip().replace("Fecha", "").replace("-", "").replace(".", "").replace(",", ".")
                                    concepto = partes[0].strip()
                                    if "/" not in valor:
                                        try:
                                            num_val = round(float(valor), 2) * (-1 if "-" in partes[1] else 1)
                                            if "ACREDITACIONES PAGO QRD" in concepto:
                                                num_val = -abs(num_val)
                                            
                                            # Sumar si el concepto ya existe en el movimiento
                                            if concepto in movimiento and isinstance(movimiento[concepto], (int, float)):
                                                movimiento[concepto] += num_val
                                            else:
                                                movimiento[concepto] = num_val
                                        except ValueError:
                                            continue
                                    else:
                                        movimiento[concepto] = partes[1]

                    if not movimientos:
                        st.error("No se encontraron liquidaciones en el PDF. Verificá el formato del archivo.")
                    else:
                        with st.spinner("Generando Excel..."):
                            df_total = pd.DataFrame(movimientos).fillna(0)
                            
                            # Integrar ACREDITACIONES PAGO QRD a QR RETENCION IIBB
                            col_acred = next((c for c in df_total.columns if "ACREDITACIONES PAGO QRD" in c), None)
                            col_qr_ret = next((c for c in df_total.columns if "QR" in c and "RETENCION" in c and "IIBB" in c), None)
                            
                            if col_acred:
                                if col_qr_ret:
                                    df_total[col_qr_ret] += df_total[col_acred]
                                else:
                                    df_total["QR RETENCION IIBB"] = df_total[col_acred]

                            columnas_qr = [col for col in df_total.columns if col.startswith("QR")]
                            df_qr = df_total[df_total[columnas_qr].sum(axis=1) != 0][["Fecha", "Liquidacion"] + columnas_qr] if columnas_qr else None
                            columnas_ajuste = [col for col in df_total.columns if "AJUSTE" in col]
                            df_ajuste = df_total[df_total[columnas_ajuste].sum(axis=1) != 0][["Fecha", "Liquidacion"] + columnas_ajuste] if columnas_ajuste else None

                            df_movimientos = df_total.drop(columns=columnas_qr + columnas_ajuste)
                            columnas_importe_neto = [col for col in df_movimientos.columns if "IMPORTE NETO" in col]
                            columnas_ventas = [col for col in df_movimientos.columns if col.startswith("VENTAS")]
                            columnas_restantes = [col for col in df_movimientos.columns if col not in columnas_importe_neto + columnas_ventas + ["Fecha", "Liquidacion"]]
                            df_movimientos = df_movimientos[["Fecha", "Liquidacion"] + columnas_restantes + columnas_importe_neto + columnas_ventas]

                            # Obtener primer numero de liquidacion
                            primer_liq = str(int(df_movimientos["Liquidacion"].iloc[0])) if len(df_movimientos) > 0 else "0"
                            encabezado_fc = f"FC {periodo_codigo}-{primer_liq}/A"

                            # CUIT First Data (hardcoded, no aparece en extractos)
                            CUIT_FIRST_DATA = "30-52221156-3"

                            output = io.BytesIO()
                            border = Border(
                                left=Side(border_style="thin"), right=Side(border_style="thin"),
                                top=Side(border_style="thin"), bottom=Side(border_style="thin")
                            )
                            money_fmt = '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)'
                            # Colores VERDES
                            header_fill = PatternFill('solid', fgColor='2E7D32')
                            zebra_fill = PatternFill('solid', fgColor='C8E6C9')
                            header_font_white = Font(bold=True, size=11, color='FFFFFF')
                            center_align = Alignment(horizontal='center', vertical='center')

                            def formatear_hoja_liq(ws, df_hoja, columnas_ignorar, titulo_encabezado=None, nombre_entidad="", mostrar_resumen=True):
                                # Ocultar líneas de cuadrícula
                                ws.sheet_view.showGridLines = False

                                # Insertar columna vacía al principio como separador visual
                                ws.insert_cols(1)
                                ws.column_dimensions['A'].width = 4  # Ancho para espaciado visual

                                # Insertar fila vacía al principio como separador visual
                                ws.insert_rows(1)

                                n_cols = len(df_hoja.columns)
                                # Offset +1 por la columna espaciadora
                                col_offset = 1
                                first_data_col = 1 + col_offset  # Columna B
                                last_data_col = n_cols + col_offset  # Última columna de datos

                                # Encabezado: filas de titulo
                                if titulo_encabezado:
                                    ws.insert_rows(2, 6)  # 5 filas de encabezado + 1 en blanco (después de fila vacía)

                                    # 4 columnas para el encabezado (B:E)
                                    merge_end = get_column_letter(first_data_col + 3)  # 4 columnas desde B

                                    # Fila 2: LIQUIDACION DE TARJETA: (tarjeta)
                                    ws.merge_cells(f'B2:{merge_end}2')
                                    ws['B2'] = f"LIQUIDACION DE TARJETA: {tipo_tarjeta_final.upper()}"
                                    ws['B2'].font = Font(bold=True, size=14, color='FFFFFF')
                                    ws['B2'].fill = header_fill
                                    ws['B2'].alignment = center_align

                                    # Fila 3: Contribuyente
                                    ws.merge_cells(f'B3:{merge_end}3')
                                    ws['B3'] = nombre_contribuyente.upper()
                                    ws['B3'].font = Font(bold=True, size=11, color='2E7D32')
                                    ws['B3'].alignment = center_align

                                    # Fila 4: Comprobante (AAMM-NroLiq/A)
                                    ws.merge_cells(f'B4:{merge_end}4')
                                    ws['B4'] = titulo_encabezado
                                    ws['B4'].font = Font(bold=True, size=11, color='2E7D32')
                                    ws['B4'].alignment = center_align

                                    # Fila 5: Entidad bancaria
                                    ws.merge_cells(f'B5:{merge_end}5')
                                    entidad_display = nombre_entidad if nombre_entidad else banco
                                    ws['B5'] = entidad_display
                                    ws['B5'].font = Font(italic=True, size=10, color='388E3C')
                                    ws['B5'].alignment = center_align

                                    # Fila 6: Periodo
                                    ws.merge_cells(f'B6:{merge_end}6')
                                    ws['B6'] = f"PERIODO: {periodo_liq.strip()}"
                                    ws['B6'].font = Font(italic=True, size=10, color='388E3C')
                                    ws['B6'].alignment = center_align

                                    # ─── Borde negro intenso externo en el encabezado (filas 2-6, cols B:merge_end) ───
                                    thick_side = Side(border_style='thick', color='000000')
                                    no_side = Side(border_style=None)
                                    merge_end_idx = first_data_col + 3
                                    for row_i in range(2, 7):
                                        for col_i in range(first_data_col, merge_end_idx + 1):
                                            cell = ws.cell(row=row_i, column=col_i)
                                            t = thick_side if row_i == 2 else no_side
                                            b = thick_side if row_i == 6 else no_side
                                            l = thick_side if col_i == first_data_col else no_side
                                            r = thick_side if col_i == merge_end_idx else no_side
                                            cell.border = Border(top=t, bottom=b, left=l, right=r)

                                    # Fila 7: en blanco (separador)
                                    data_header_row = 8
                                    data_start_row = 9
                                else:
                                    data_header_row = 2
                                    data_start_row = 3

                                # Ajustar columnas_ignorar con offset (B=col2, C=col3, etc.)
                                columnas_ignorar_offset = [get_column_letter(ord(c) - ord('A') + 1 + col_offset) for c in columnas_ignorar]

                                # Estilo de encabezados de columna
                                for col_idx in range(first_data_col, last_data_col + 1):
                                    cell = ws.cell(row=data_header_row, column=col_idx)
                                    cell.font = header_font_white
                                    cell.fill = header_fill
                                    cell.alignment = center_align
                                    cell.border = border

                                # Estilo de datos (sin bordes internos)
                                last_data_row = data_start_row + len(df_hoja) - 1
                                for row_idx in range(data_start_row, last_data_row + 1):
                                    for col_idx in range(first_data_col, last_data_col + 1):
                                        cell = ws.cell(row=row_idx, column=col_idx)
                                        cell.alignment = center_align
                                        if cell.column_letter not in columnas_ignorar_offset:
                                            if isinstance(cell.value, (int, float)):
                                                cell.number_format = money_fmt
                                    # Zebra verde
                                    if (row_idx - data_start_row) % 2 == 0:
                                        for col_idx in range(first_data_col, last_data_col + 1):
                                            ws.cell(row=row_idx, column=col_idx).fill = zebra_fill

                                # Fila TOTAL
                                total_row = last_data_row + 1
                                fc = first_data_col
                                fc_letter = get_column_letter(fc)
                                fc1_letter = get_column_letter(fc + 1)
                                ws.merge_cells(f'{fc_letter}{total_row}:{fc1_letter}{total_row}')
                                ws[f'{fc_letter}{total_row}'] = "TOTAL"
                                ws[f'{fc_letter}{total_row}'].font = Font(bold=True, size=11, color='FFFFFF')
                                ws[f'{fc_letter}{total_row}'].fill = header_fill
                                ws[f'{fc_letter}{total_row}'].alignment = center_align
                                ws.cell(row=total_row, column=fc + 1).fill = header_fill

                                for col_idx in range(fc + 2, last_data_col + 1):
                                    cell = ws.cell(row=total_row, column=col_idx)
                                    col_letter = get_column_letter(col_idx)
                                    cell.value = f"=SUM({col_letter}{data_start_row}:{col_letter}{last_data_row})"
                                    cell.number_format = money_fmt
                                    cell.font = Font(bold=True, size=10, color='FFFFFF')
                                    cell.fill = header_fill
                                    cell.alignment = center_align

                                # ─── Borde negro intenso externo + verticales en header y total ───
                                thick_side = Side(border_style='thick', color='000000')
                                no_side = Side(border_style=None)
                                for row_i in range(data_header_row, total_row + 1):
                                    is_header = (row_i == data_header_row)
                                    is_total = (row_i == total_row)
                                    is_special_row = is_header or is_total
                                    for col_i in range(first_data_col, last_data_col + 1):
                                        cell = ws.cell(row=row_i, column=col_i)
                                        t = thick_side if is_header else (thick_side if is_total else no_side)
                                        b = thick_side if is_total else (thick_side if is_header else no_side)
                                        # Verticales thick en header y total, solo extremos en datos
                                        l = thick_side if (col_i == first_data_col or is_special_row) else no_side
                                        r = thick_side if (col_i == last_data_col or is_special_row) else no_side
                                        cell.border = Border(top=t, bottom=b, left=l, right=r)

                                # Auto-ajustar columnas (solo las de datos)
                                for col_idx in range(first_data_col, last_data_col + 1):
                                    col_letter = get_column_letter(col_idx)
                                    max_len = max(
                                        len(str(ws.cell(row=r, column=col_idx).value or ''))
                                        for r in range(data_header_row, min(total_row + 1, data_header_row + 50))
                                    )
                                    ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

                                # ─── CARGO TERMINAL y ACREDITACIONES highlight (siempre) ───────────────────
                                cols_hoja = list(df_hoja.columns)
                                cargo_terminal_cols = []
                                acreditaciones_cols = []
                                for i, col in enumerate(cols_hoja):
                                    real_col = i + 1 + col_offset  # +1 por offset de columna espaciadora
                                    if "CARGO TERMINAL" in col.upper():
                                        cargo_terminal_cols.append((real_col, get_column_letter(real_col)))
                                    if "ACREDITACIONES PAGO QRD" in col.upper():
                                        acreditaciones_cols.append((real_col, get_column_letter(real_col)))

                                yellow_fill = PatternFill('solid', fgColor='FFD600')
                                for col_idx_ct, col_letter_ct in cargo_terminal_cols:
                                    hdr_cell = ws.cell(row=data_header_row, column=col_idx_ct)
                                    hdr_cell.fill = PatternFill('solid', fgColor='FF6F00')
                                    hdr_cell.font = Font(bold=True, size=11, color='FFFFFF')
                                    from openpyxl.comments import Comment
                                    hdr_cell.comment = Comment("ATENCION: Cargo terminal detectado", "Sistema")
                                    for row_idx in range(data_start_row, total_row + 1):
                                        ws.cell(row=row_idx, column=col_idx_ct).fill = yellow_fill

                                for col_idx_ac, col_letter_ac in acreditaciones_cols:
                                    hdr_cell = ws.cell(row=data_header_row, column=col_idx_ac)
                                    hdr_cell.fill = yellow_fill
                                    hdr_cell.font = Font(bold=True, size=11, color='000000')
                                    for row_idx in range(data_start_row, total_row + 1):
                                        ws.cell(row=row_idx, column=col_idx_ac).fill = yellow_fill

                                # ─── Tabla resumen de impuestos (solo Liquidaciones) ──────
                                if mostrar_resumen:
                                    iva21_col_letters = []
                                    iva105_col_letters = []
                                    perc_iva_col_letters = []
                                    perc_iibb_col_letters = []
                                    sirtac_col_letters = []

                                    for i, col in enumerate(cols_hoja):
                                        col_upper = col.upper()
                                        col_letter = get_column_letter(i + 1 + col_offset)
                                        if "IVA" in col_upper and not col_upper.startswith("PER"):
                                            if "IVA RI" in col_upper or ("10,50" not in col and "10.50" not in col):
                                                iva21_col_letters.append(col_letter)
                                            else:
                                                iva105_col_letters.append(col_letter)
                                        if col_upper.startswith("PER") and "IVA" in col_upper:
                                            perc_iva_col_letters.append(col_letter)
                                        elif col_upper.startswith("PER") and "IVA" not in col_upper:
                                            perc_iibb_col_letters.append(col_letter)
                                        if "SIRTAC" in col_upper:
                                            sirtac_col_letters.append(col_letter)

                                    resumen_items = []
                                    tr = total_row

                                    if iva21_col_letters:
                                        if len(iva21_col_letters) == 1:
                                            iva_ref = f"{iva21_col_letters[0]}{tr}"
                                        else:
                                            iva_ref = "+".join(f"{cl}{tr}" for cl in iva21_col_letters)
                                        resumen_items.append(("NETO 21", f"=ABS({iva_ref})/0.21"))
                                        resumen_items.append(("IVA 21", f"=ABS({iva_ref})"))

                                    if iva105_col_letters:
                                        if len(iva105_col_letters) == 1:
                                            iva105_ref = f"{iva105_col_letters[0]}{tr}"
                                        else:
                                            iva105_ref = "+".join(f"{cl}{tr}" for cl in iva105_col_letters)
                                        resumen_items.append(("NETO 10.5", f"=ABS({iva105_ref})/0.105"))
                                        resumen_items.append(("IVA 10.5", f"=ABS({iva105_ref})"))

                                    if perc_iva_col_letters:
                                        if len(perc_iva_col_letters) == 1:
                                            p_ref = f"{perc_iva_col_letters[0]}{tr}"
                                        else:
                                            p_ref = "+".join(f"{cl}{tr}" for cl in perc_iva_col_letters)
                                        resumen_items.append(("PERC. IVA", f"=ABS({p_ref})"))

                                    if sirtac_col_letters:
                                        if len(sirtac_col_letters) == 1:
                                            s_ref = f"{sirtac_col_letters[0]}{tr}"
                                        else:
                                            s_ref = "+".join(f"{cl}{tr}" for cl in sirtac_col_letters)
                                        resumen_items.append(("SIRTAC", f"=ABS({s_ref})"))

                                    if perc_iibb_col_letters:
                                        if len(perc_iibb_col_letters) == 1:
                                            pi_ref = f"{perc_iibb_col_letters[0]}{tr}"
                                        else:
                                            pi_ref = "+".join(f"{cl}{tr}" for cl in perc_iibb_col_letters)
                                        resumen_items.append(("PERC. IIBB", f"=ABS({pi_ref})"))

                                    if resumen_items:
                                        resumen_start = total_row + 2
                                        # Resumen con 3 columnas (B:D)
                                        res_start_letter = get_column_letter(first_data_col)
                                        res_end_letter = get_column_letter(first_data_col + 2)
                                        ws.merge_cells(f'{res_start_letter}{resumen_start}:{res_end_letter}{resumen_start}')
                                        ws[f'{res_start_letter}{resumen_start}'] = "RESUMEN IMPOSITIVO"
                                        ws[f'{res_start_letter}{resumen_start}'].font = Font(bold=True, size=11, color='FFFFFF')
                                        ws[f'{res_start_letter}{resumen_start}'].fill = header_fill
                                        ws[f'{res_start_letter}{resumen_start}'].alignment = center_align

                                        for idx, (concepto, formula) in enumerate(resumen_items):
                                            r = resumen_start + 1 + idx
                                            merge_a = get_column_letter(first_data_col)
                                            merge_b = get_column_letter(first_data_col + 1)  # Concepto ocupa 2 cols (B:C)
                                            val_col = first_data_col + 2  # Valor en col D
                                            ws.merge_cells(f'{merge_a}{r}:{merge_b}{r}')
                                            ws[f'{merge_a}{r}'] = concepto
                                            ws[f'{merge_a}{r}'].font = Font(bold=True, size=10)
                                            ws[f'{merge_a}{r}'].alignment = center_align
                                            cell_val = ws.cell(row=r, column=val_col)
                                            cell_val.value = formula
                                            cell_val.number_format = money_fmt
                                            cell_val.alignment = center_align
                                            if idx % 2 == 0:
                                                ws[f'{merge_a}{r}'].fill = zebra_fill
                                                ws.cell(row=r, column=first_data_col + 1).fill = zebra_fill
                                                cell_val.fill = zebra_fill

                                        # Fila TOTAL del resumen
                                        r_total = resumen_start + 1 + len(resumen_items)
                                        first_val_row = resumen_start + 1
                                        last_val_row = r_total - 1
                                        merge_a = get_column_letter(first_data_col)
                                        merge_b = get_column_letter(first_data_col + 1)
                                        val_col = first_data_col + 2
                                        ws.merge_cells(f'{merge_a}{r_total}:{merge_b}{r_total}')
                                        ws[f'{merge_a}{r_total}'] = "TOTAL"
                                        ws[f'{merge_a}{r_total}'].font = Font(bold=True, size=11, color='FFFFFF')
                                        ws[f'{merge_a}{r_total}'].fill = header_fill
                                        ws[f'{merge_a}{r_total}'].alignment = center_align
                                        ws.cell(row=r_total, column=first_data_col + 1).fill = header_fill
                                        val_col_letter = get_column_letter(val_col)
                                        cell_total = ws.cell(row=r_total, column=val_col)
                                        cell_total.value = f"=SUM({val_col_letter}{first_val_row}:{val_col_letter}{last_val_row})"
                                        cell_total.number_format = money_fmt
                                        cell_total.font = Font(bold=True, size=10, color='FFFFFF')
                                        cell_total.fill = header_fill
                                        cell_total.alignment = center_align

                                        # ─── Borde negro intenso externo en resumen impositivo ───
                                        thick_side = Side(border_style='thick', color='000000')
                                        no_side = Side(border_style=None)
                                        res_first_col = first_data_col
                                        res_last_col = first_data_col + 2  # 3 columnas
                                        for row_i in range(resumen_start, r_total + 1):
                                            for col_i in range(res_first_col, res_last_col + 1):
                                                cell = ws.cell(row=row_i, column=col_i)
                                                t = thick_side if row_i == resumen_start else no_side
                                                b = thick_side if row_i == r_total else no_side
                                                l = thick_side if col_i == res_first_col else no_side
                                                r_s = thick_side if col_i == res_last_col else no_side
                                                cell.border = Border(top=t, bottom=b, left=l, right=r_s)

                            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                                df_movimientos.to_excel(writer, sheet_name="Liquidaciones", index=False)
                                if df_qr is not None:
                                    df_qr.to_excel(writer, sheet_name="QR", index=False)
                                if df_ajuste is not None:
                                    df_ajuste.to_excel(writer, sheet_name="AJUSTE", index=False)

                                # Formatear hojas
                                wb = writer.book
                                formatear_hoja_liq(wb["Liquidaciones"], df_movimientos, ["A", "B"], encabezado_fc, banco, mostrar_resumen=True)
                                if df_qr is not None:
                                    formatear_hoja_liq(wb["QR"], df_qr, ["A", "B"], encabezado_fc, "First Data", mostrar_resumen=False)
                                if df_ajuste is not None:
                                    formatear_hoja_liq(wb["AJUSTE"], df_ajuste, ["A", "B"], encabezado_fc, banco, mostrar_resumen=False)

                            output.seek(0)

                        st.success("✓  Liquidaciones procesadas con éxito")

                        # Stats
                        st.markdown(f"""
                        <div class="stats-row">
                            <div class="stat-chip">
                                <span class="stat-val">{len(movimientos)}</span>
                                <span class="stat-lbl">Liquidaciones</span>
                            </div>
                            <div class="stat-chip">
                                <span class="stat-val">{len(columnas_qr)}</span>
                                <span class="stat-lbl">Cols. QR</span>
                            </div>
                            <div class="stat-chip">
                                <span class="stat-val">{len(columnas_ajuste)}</span>
                                <span class="stat-lbl">Cols. Ajuste</span>
                            </div>
                        </div>
                        """, unsafe_allow_html=True)

                        st.info(
                            f"**{nombre_contribuyente}** · {tipo_tarjeta_final} · "
                            f"**{encabezado_fc}** · {banco} · "
                            f"{len(reader.pages)} páginas"
                        )

                        st.download_button(
                            label="↓  Descargar Excel de Liquidaciones",
                            data=output,
                            file_name=f"{liq_filename}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True,
                        )

                except Exception as e:
                    st.error(f"Error al procesar el archivo: {str(e)}")
                    st.exception(e)

        st.markdown('</div>', unsafe_allow_html=True)

    else:
        st.markdown("""
        <div style="
            text-align: center;
            padding: 2rem 1rem;
            font-family: 'Space Mono', monospace;
            font-size: 0.72rem;
            color: #6b7280;
            letter-spacing: 0.12em;
        ">
            ESPERANDO ARCHIVO PDF · PASO 01
        </div>
        """, unsafe_allow_html=True)


elif herramienta == TOOL_DEDUCCIONES:
    # ───────────────────────────────────────────────────────────────────────────────
    # HERRAMIENTA: Limpieza Excel Deducciones IVA/Ganancias
    # ───────────────────────────────────────────────────────────────────────────────
    st.markdown('<div class="card"><div class="card-label">01 · Archivo de Deducciones (.xls / .xlsx)</div>', unsafe_allow_html=True)
    uploaded_ded = st.file_uploader(
        "Subí el Excel descargado de Mis Retenciones/Percepciones de ARCA",
        type=["xls", "xlsx"],
        label_visibility="visible",
        key="deducciones_xls"
    )
    st.markdown('</div>', unsafe_allow_html=True)

    if uploaded_ded:
        st.success(f"**{uploaded_ded.name}** listo para procesar")

        st.markdown('<div class="card"><div class="card-label">02 · Datos del contribuyente</div>', unsafe_allow_html=True)
        nombre_ded = st.text_input("Nombre / Razón Social del contribuyente", value="", key="nombre_deducciones")
        st.markdown('</div>', unsafe_allow_html=True)

        st.markdown('<div class="card"><div class="card-label">03 · Procesar</div>', unsafe_allow_html=True)

        if st.button("⬡  Limpiar y Estilizar"):
            if not nombre_ded.strip():
                st.error("Ingresá el nombre del contribuyente para continuar.")
            else:
              try:
                with st.spinner("Procesando Excel de deducciones..."):
                    # Leer el archivo
                    df_ded = pd.read_excel(io.BytesIO(uploaded_ded.getvalue()))

                    if df_ded.empty:
                        st.error("El archivo está vacío.")
                        st.stop()

                    # ── Detectar tipo de impuesto ──
                    desc_imp_col = None
                    for c in df_ded.columns:
                        if 'descripci' in c.lower() and 'impuesto' in c.lower():
                            desc_imp_col = c
                            break
                    tipo_deduccion = 'DEDUCCIONES'
                    if desc_imp_col and not df_ded[desc_imp_col].dropna().empty:
                        primer_imp = str(df_ded[desc_imp_col].dropna().iloc[0]).upper()
                        if 'GANANCIA' in primer_imp:
                            tipo_deduccion = 'DEDUCCIONES GANANCIAS'
                        elif 'VALOR AGRE' in primer_imp or 'IVA' in primer_imp:
                            tipo_deduccion = 'DEDUCCIONES IVA'
                        elif 'SIRE' in primer_imp:
                            tipo_deduccion = 'SIRE IVA'

                    # ── Eliminar columnas vacías y redundantes ──
                    cols_drop = []
                    for c in df_ded.columns:
                        cl = c.lower()
                        if c.strip() == 'Impuesto' or c.strip() == 'Régimen':
                            cols_drop.append(c)
                    df_ded = df_ded.drop(columns=[c for c in cols_drop if c in df_ded.columns], errors='ignore')

                    # ── Renombrar columnas ──
                    # Detectar y renombrar la columna de Razón Social dinámicamente
                    for c in df_ded.columns:
                        cl = c.lower()
                        if 'denominaci' in cl and 'raz' in cl:
                            df_ded = df_ded.rename(columns={c: 'Razón Social'})
                            break

                    RENAME_DED = {
                        'CUIT Agente Ret./Perc.': 'CUIT',
                        'Descripción Impuesto': 'Impuesto',
                        'Descripción Régimen': 'Régimen',
                        'Fecha Ret./Perc.': 'Fecha',
                        'Número Certificado': 'Nro. Certificado',
                        'Descripción Operación': 'Operación',
                        'Importe Ret./Perc.': 'Importe',
                        'Número Comprobante': 'Nro. Comprobante',
                        'Fecha Comprobante': 'Fecha Comp.',
                        'Descripción Comprobante': 'Comprobante',
                        'Fecha Registración DJ Ag.Ret.': 'Fecha Reg. DJ',
                    }
                    df_ded = df_ded.rename(columns=RENAME_DED)

                    # ── Formatear CUIT como XX-XXXXXXXX-X ──
                    if 'CUIT' in df_ded.columns:
                        def format_cuit(val):
                            s = str(int(val)) if not pd.isna(val) else ''
                            if len(s) == 11:
                                return f"{s[:2]}-{s[2:10]}-{s[10]}"
                            return s
                        df_ded['CUIT'] = df_ded['CUIT'].apply(format_cuit)

                    # ── Ordenar por Fecha ascendente ──
                    if 'Fecha' in df_ded.columns:
                        try:
                            df_ded['_fecha_sort'] = pd.to_datetime(df_ded['Fecha'], format='%d/%m/%Y', errors='coerce')
                            df_ded = df_ded.sort_values('_fecha_sort', ascending=True).drop(columns=['_fecha_sort'])
                        except Exception:
                            pass

                    # ── Separar Retenciones y Percepciones ──
                    op_col = 'Operación'
                    df_ret = df_ded[df_ded[op_col].str.upper().str.contains('RETENCION', na=False)].copy() if op_col in df_ded.columns else pd.DataFrame()
                    df_per = df_ded[df_ded[op_col].str.upper().str.contains('PERCEPCION', na=False)].copy() if op_col in df_ded.columns else pd.DataFrame()
                    # Si no hay columna Operación, todo va a una hoja genérica
                    if op_col not in df_ded.columns:
                        df_ret = df_ded
                        df_per = pd.DataFrame()

                    # Eliminar columnas Impuesto y Operación (ya discriminadas por hoja)
                    for df_part in [df_ret, df_per]:
                        for drop_c in ['Impuesto', 'Operación']:
                            if drop_c in df_part.columns:
                                df_part.drop(columns=[drop_c], inplace=True)

                    # Mover Importe al final
                    for df_part in [df_ret, df_per]:
                        if 'Importe' in df_part.columns:
                            imp_data = df_part.pop('Importe')
                            df_part['Importe'] = imp_data

                    # ── Estilos dorados/ámbar ──
                    title_font = Font(bold=True, size=14, color='FFFFFF')
                    title_fill = PatternFill('solid', fgColor='BF8F00')
                    header_font = Font(bold=True, size=10, color='FFFFFF')
                    header_fill = PatternFill('solid', fgColor='D4A017')
                    header_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
                    center_align = Alignment(horizontal='center', vertical='center')
                    thin_border = Border(
                        left=Side(style='thin'), right=Side(style='thin'),
                        top=Side(style='thin'), bottom=Side(style='thin')
                    )
                    zebra_fill = PatternFill('solid', fgColor='FFF2CC')
                    accounting_fmt = '_-"$"* #,##0.00_-;-"$"* #,##0.00_-;_-"$"* "-"??_-;_-@_-'

                    def _style_ded_sheet(ws, df_sheet, sheet_title, n_r, n_c, col_list):
                        """Aplica estilos dorados a una hoja de deducciones."""
                        # Fila 1: Nombre contribuyente
                        ws.merge_cells(f'A1:{get_column_letter(n_c)}1')
                        ws['A1'] = nombre_ded.strip().upper()
                        ws['A1'].font = title_font
                        ws['A1'].fill = title_fill
                        ws['A1'].alignment = center_align

                        # Fila 2: Título hoja + tipo y cantidad
                        ws.merge_cells(f'A2:{get_column_letter(n_c)}2')
                        ws['A2'] = f'{sheet_title} — {tipo_deduccion} — {n_r} registros'
                        ws['A2'].font = Font(italic=True, size=10, color='BF8F00')
                        ws['A2'].alignment = center_align

                        # Encabezados (fila 6)
                        for ci in range(1, n_c + 1):
                            cell = ws.cell(row=6, column=ci)
                            cell.font = header_font
                            cell.fill = header_fill
                            cell.alignment = header_align
                            cell.border = thin_border

                        # Columna Importe
                        imp_idx = col_list.index('Importe') + 1 if 'Importe' in col_list else None

                        # Datos (fila 7+)
                        for ri in range(7, n_r + 7):
                            for ci in range(1, n_c + 1):
                                cell = ws.cell(row=ri, column=ci)
                                cell.alignment = center_align
                                cell.border = thin_border
                                if ci == imp_idx:
                                    cell.number_format = accounting_fmt
                            if (ri - 7) % 2 == 0:
                                for ci in range(1, n_c + 1):
                                    ws.cell(row=ri, column=ci).fill = zebra_fill

                        # Fila TOTAL
                        if imp_idx:
                            tr = n_r + 7
                            ws.merge_cells(f'A{tr}:{get_column_letter(imp_idx - 1)}{tr}')
                            ws[f'A{tr}'] = 'TOTAL'
                            ws[f'A{tr}'].font = Font(bold=True)
                            ws[f'A{tr}'].alignment = Alignment(horizontal='right')
                            il = get_column_letter(imp_idx)
                            tc = ws.cell(row=tr, column=imp_idx)
                            tc.value = f'=SUM({il}7:{il}{tr - 1})'
                            tc.font = Font(bold=True)
                            tc.border = Border(top=Side(style='double'))
                            tc.number_format = accounting_fmt
                            tc.alignment = center_align

                        # Autofit
                        for ci in range(1, n_c + 1):
                            cl = get_column_letter(ci)
                            mx = len(str(ws.cell(row=6, column=ci).value or ''))
                            for ri in range(7, min(n_r + 7, 57)):
                                v = ws.cell(row=ri, column=ci).value
                                if v:
                                    mx = max(mx, len(str(v)))
                            ws.column_dimensions[cl].width = min(mx + 3, 45)

                    # ── Generar Excel ──
                    output = io.BytesIO()
                    sheets_written = []

                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        if not df_ret.empty:
                            df_ret.to_excel(writer, sheet_name='Retenciones', index=False, startrow=5)
                            ws_ret = writer.sheets['Retenciones']
                            ret_cols = list(df_ret.columns)
                            _style_ded_sheet(ws_ret, df_ret, 'RETENCIONES', len(df_ret), len(ret_cols), ret_cols)
                            sheets_written.append(('Retenciones', len(df_ret)))

                        if not df_per.empty:
                            df_per.to_excel(writer, sheet_name='Percepciones', index=False, startrow=5)
                            ws_per = writer.sheets['Percepciones']
                            per_cols = list(df_per.columns)
                            _style_ded_sheet(ws_per, df_per, 'PERCEPCIONES', len(df_per), len(per_cols), per_cols)
                            sheets_written.append(('Percepciones', len(df_per)))

                    output.seek(0)
                    n_rows = len(df_ded)

                st.success("✓  Proceso completado con éxito")

                # Stats
                stats_html = f'<div class="stats-row"><div class="stat-chip"><span class="stat-val">{n_rows}</span><span class="stat-lbl">Total</span></div>'
                for sname, scount in sheets_written:
                    stats_html += f'<div class="stat-chip"><span class="stat-val">{scount}</span><span class="stat-lbl">{sname}</span></div>'
                stats_html += '</div>'
                st.markdown(stats_html, unsafe_allow_html=True)

                ded_filename = f"{Path(uploaded_ded.name).stem}_limpio.xlsx"
                st.download_button(
                    label="↓  Descargar Excel Limpio",
                    data=output,
                    file_name=ded_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

              except Exception as e:
                st.error(f"Error al procesar el archivo: {str(e)}")
                st.exception(e)

        st.markdown('</div>', unsafe_allow_html=True)

    else:
        st.markdown("""
        <div style="
            text-align: center;
            padding: 2rem 1rem;
            font-family: 'Space Mono', monospace;
            font-size: 0.72rem;
            color: #6b7280;
            letter-spacing: 0.12em;
        ">
            ESPERANDO ARCHIVO EXCEL · PASO 01
        </div>
        """, unsafe_allow_html=True)


elif herramienta == TOOL_ARBA:
    # ───────────────────────────────────────────────────────────────────────────────
    # HERRAMIENTA: Archivo Percepciones ARBA
    # ───────────────────────────────────────────────────────────────────────────────
    st.markdown('<div class="card"><div class="card-label">01 · Archivo fuente (Movimientos Ventas)</div>', unsafe_allow_html=True)
    uploaded_arba = st.file_uploader(
        "Arrastrá tu archivo de movimientos de ventas o hacé click para seleccionarlo",
        type=["txt", "prn"],
        label_visibility="visible",
        key="arba_file"
    )
    st.markdown('</div>', unsafe_allow_html=True)

    if uploaded_arba:
        arba_filename = Path(uploaded_arba.name).stem
        st.success(f"**{uploaded_arba.name}** listo para procesar")

        # ─── Card 02: Periodo ──────────────────────────────────────────────────────
        st.markdown('<div class="card"><div class="card-label">02 · Periodo</div>', unsafe_allow_html=True)
        periodo_arba = st.text_input(
            "Ingresá el periodo (MM/AAAA)",
            value="",
            placeholder="Ej: 03/2026",
            key="arba_periodo"
        )
        st.markdown('</div>', unsafe_allow_html=True)

        # ─── Card 03: Procesar ─────────────────────────────────────────────────────
        st.markdown('<div class="card"><div class="card-label">03 · Generar TXT Percepciones ARBA</div>', unsafe_allow_html=True)

        if st.button("⬡  Generar archivo ARBA"):
            # Validar periodo obligatorio
            periodo_limpio = periodo_arba.strip()
            periodo_match = re.match(r'^(\d{2})/(\d{4})$', periodo_limpio)
            if not periodo_match:
                st.error("El periodo es obligatorio y debe tener el formato **MM/AAAA** (ej: 03/2026)")
            else:
                mes_p = periodo_match.group(1)
                anio_p = periodo_match.group(2)
                try:
                    with st.spinner("Procesando movimientos de ventas..."):
                        raw_bytes = uploaded_arba.getvalue()
                        content_str = raw_bytes.decode('latin-1', errors='replace')
                        movimientos, metadata = parsear_archivo(content=content_str)

                        # Inyectar periodo ingresado por el usuario
                        metadata['periodo'] = f"Desde el 01/{mes_p}/{anio_p} hasta el 28/{mes_p}/{anio_p}"

                        txt_arba = generar_percepciones_arba_txt(movimientos, metadata)

                    if not txt_arba.strip():
                        st.warning("No se encontraron percepciones IIBB Buenos Aires en los movimientos.")
                    else:
                        st.success("✓  Archivo Percepciones ARBA generado con éxito")

                        # Stats
                        n_lineas = len(txt_arba.splitlines())
                        st.markdown(f"""
                        <div class="stats-row">
                            <div class="stat-chip">
                                <span class="stat-val">{len(movimientos)}</span>
                                <span class="stat-lbl">Movimientos</span>
                            </div>
                            <div class="stat-chip">
                                <span class="stat-val">{n_lineas}</span>
                                <span class="stat-lbl">Líneas TXT</span>
                            </div>
                        </div>
                        """, unsafe_allow_html=True)

                        # Nombre archivo ARBA: AR-CUIT-PERIODO-P7-LOTE.txt
                        cuit_empresa = metadata.get('cuit_empresa', '').replace('-', '')
                        periodo_file = f"{anio_p}{mes_p}"  # YYYYMM
                        arba_download_name = f"AR-{cuit_empresa}-{periodo_file}-P7-1.txt"

                        st.download_button(
                            label=f"↓  Descargar TXT ({arba_download_name})",
                            data=txt_arba.encode("latin-1", errors="replace"),
                            file_name=arba_download_name,
                            mime="text/plain",
                            use_container_width=True,
                        )

                except Exception as e:
                    st.error(f"Error al procesar el archivo: {str(e)}")
                    st.exception(e)

        st.markdown('</div>', unsafe_allow_html=True)

    else:
        st.markdown("""
        <div style="
            text-align: center;
            padding: 2rem 1rem;
            font-family: 'Space Mono', monospace;
            font-size: 0.72rem;
            color: #6b7280;
            letter-spacing: 0.12em;
        ">
            ESPERANDO ARCHIVO · PASO 01
        </div>
        """, unsafe_allow_html=True)


elif herramienta == TOOL_CRUCE_CONCEPTO:
    # ───────────────────────────────────────────────────────────────────────────────
    # HERRAMIENTA: Cruce Concepto (TXT + Excel Sistema)
    # ───────────────────────────────────────────────────────────────────────────────
    st.markdown('<div class="card"><div class="card-label">01 · Archivo TXT (Comprobantes)</div>', unsafe_allow_html=True)
    uploaded_txt_concepto = st.file_uploader(
        "Subí el .txt de Comprobantes de Compras (del sistema)",
        type=["txt", "prn"],
        label_visibility="visible",
        key="cruce_concepto_txt"
    )
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="card"><div class="card-label">02 · Archivo Excel Sistema (.xls)</div>', unsafe_allow_html=True)
    uploaded_xls_concepto = st.file_uploader(
        "Subí el Excel del sistema (.xls) con las compras",
        type=["xls", "xlsx"],
        label_visibility="visible",
        key="cruce_concepto_xls"
    )
    st.markdown('</div>', unsafe_allow_html=True)

    if uploaded_txt_concepto and uploaded_xls_concepto:
        st.success(f"**{uploaded_txt_concepto.name}** + **{uploaded_xls_concepto.name}** listos para cruzar")

        st.markdown('<div class="card"><div class="card-label">03 · Procesar</div>', unsafe_allow_html=True)

        if st.button("⬡  Cruzar Concepto"):
            try:
                with st.spinner("Parseando TXT..."):
                    txt_content = uploaded_txt_concepto.getvalue().decode("latin-1")
                    transacciones, meta_txt = parsear_archivo(content=txt_content)

                if not transacciones:
                    st.error("No se encontraron transacciones en el TXT. Verificá el formato.")
                else:
                    with st.spinner("Leyendo Excel Sistema..."):
                        df_xls = pd.read_excel(io.BytesIO(uploaded_xls_concepto.getvalue()))

                        # Limpiar: quitar filas completamente vacías y fila de TOTALES
                        df_xls = df_xls.dropna(how='all')
                        # Buscar TOTALES en Fecha, Nombre y Cond (donde suele estar)
                        for col_check in ['Fecha', 'Nombre', 'Cond']:
                            if col_check in df_xls.columns:
                                df_xls = df_xls[~df_xls[col_check].astype(str).str.upper().str.contains('TOTAL', na=False)]
                        df_xls = df_xls.reset_index(drop=True)

                    # ── Construir lookup de Concepto desde TXT ──────────────
                    # Clave: tipo + pv + nro (sin letra) + cuit (sin guiones)
                    concepto_lookup = {}
                    for t in transacciones:
                        numero_raw = t['Numero']
                        pv_txt = numero_raw.split('-')[0] if '-' in numero_raw else numero_raw[:5]
                        resto_num = numero_raw.split('-')[1] if '-' in numero_raw else numero_raw[5:]
                        # Quitar letra final del nro
                        nro_txt = resto_num[:-1] if resto_num and resto_num[-1].isalpha() else resto_num
                        cuit_txt = t['CUIT'].replace('-', '')

                        # Normalizar: quitar ceros a la izquierda del PV y Nro
                        try:
                            pv_norm = str(int(pv_txt))
                        except ValueError:
                            pv_norm = pv_txt
                        try:
                            nro_norm = str(int(nro_txt))
                        except ValueError:
                            nro_norm = nro_txt

                        key = f"{t['Tipo']}|{pv_norm}|{nro_norm}|{cuit_txt}"
                        concepto_lookup[key] = (t['Concepto'], t['Letra'])

                    # ── Parsear columnas del Excel ──────────────────────────
                    # El Excel tiene: Fecha, TC, Numero, Nombre, Cond, C.U.I.T., ...
                    # Numero tiene formato "PPPPP-NNNNNNNNNN/L" o similar
                    def extraer_key_xls(row):
                        tc = str(row.get('TC', '')).strip()
                        numero = str(row.get('Numero', '')).strip()
                        cuit = str(row.get('C.U.I.T.', '')).replace('-', '').replace('.', '').strip()

                        # Separar PV y Nro del campo Numero (ej: 00003-00021793/A)
                        if '-' in numero:
                            parts = numero.split('-', 1)
                            pv_raw = parts[0]
                            nro_raw = parts[1]
                        elif '/' in numero:
                            parts = numero.split('/', 1)
                            pv_raw = parts[0]
                            nro_raw = parts[1]
                        else:
                            pv_raw = numero[:5] if len(numero) >= 5 else numero
                            nro_raw = numero[5:] if len(numero) > 5 else ''

                        # Quitar letra y / del nro
                        nro_clean = re.sub(r'[/A-Za-z]+$', '', nro_raw).strip()

                        try:
                            pv_norm = str(int(pv_raw))
                        except ValueError:
                            pv_norm = pv_raw
                        try:
                            nro_norm = str(int(nro_clean))
                        except ValueError:
                            nro_norm = nro_clean

                        return f"{tc}|{pv_norm}|{nro_norm}|{cuit}"

                    with st.spinner("Cruzando datos..."):
                        # Filas "cabecera" son las que tienen Fecha; las demás son sub-filas del mismo movimiento
                        mask_header = df_xls['Fecha'].notna() & (df_xls['Fecha'] != '')

                        # Agregar columnas de Concepto y Jurisdicción
                        conceptos_cod = []
                        jurisdicciones = []
                        matched = 0
                        last_concepto = ''
                        last_jur = ''
                        for idx, row in df_xls.iterrows():
                            if mask_header.iloc[idx]:
                                # Es fila cabecera → buscar concepto y jurisdicción
                                key = extraer_key_xls(row)
                                result = concepto_lookup.get(key)
                                if result is not None:
                                    matched += 1
                                    last_concepto, last_jur = result
                                else:
                                    last_concepto = ''
                                    last_jur = ''
                                conceptos_cod.append(last_concepto)
                                jurisdicciones.append(last_jur)
                            else:
                                # Sub-fila → propagar de la cabecera
                                conceptos_cod.append(last_concepto)
                                jurisdicciones.append(last_jur)

                        # Insertar Concepto y Jurisdicción después de C.U.I.T.
                        cuit_pos = df_xls.columns.get_loc('C.U.I.T.') + 1 if 'C.U.I.T.' in df_xls.columns else len(df_xls.columns)
                        df_xls.insert(cuit_pos, 'Concepto', conceptos_cod)
                        df_xls.insert(cuit_pos + 1, 'Jur.', jurisdicciones)

                        # Forward-fill columnas identificatorias a sub-filas
                        for col_ff in ['Fecha', 'TC', 'Numero', 'Nombre', 'Cond', 'C.U.I.T.', 'Concepto', 'Jur.']:
                            if col_ff in df_xls.columns:
                                df_xls[col_ff] = df_xls[col_ff].ffill()

                        # Separar Fecha (dd/mm/yyyy) en Dia, Mes, Año
                        if 'Fecha' in df_xls.columns:
                            fecha_pos = df_xls.columns.get_loc('Fecha')
                            fecha_str = df_xls['Fecha'].astype(str)
                            # Intentar parsear como dd/mm/yyyy
                            partes_fecha = fecha_str.str.split('/', expand=True)
                            if partes_fecha.shape[1] >= 3:
                                df_xls.insert(fecha_pos, 'Dia', pd.to_numeric(partes_fecha[0], errors='coerce').fillna(0).astype(int))
                                df_xls.insert(fecha_pos + 1, 'Mes', pd.to_numeric(partes_fecha[1], errors='coerce').fillna(0).astype(int))
                                df_xls.insert(fecha_pos + 2, 'Año', pd.to_numeric(partes_fecha[2], errors='coerce').fillna(0).astype(int))
                                df_xls.drop(columns=['Fecha'], inplace=True)

                        if 'Numero' in df_xls.columns:
                            num_pos = df_xls.columns.get_loc('Numero')
                            def split_numero(val):
                                s = str(val).strip()
                                # Quitar /Letra o Letra final del numero
                                letra = ''
                                if '/' in s:
                                    parts = s.rsplit('/', 1)
                                    s = parts[0]
                                    letra = parts[1] if len(parts) > 1 else ''
                                elif s and s[-1].isalpha():
                                    letra = s[-1]
                                    s = s[:-1]
                                return s, letra

                            numero_list, letra_list = [], []
                            for val in df_xls['Numero']:
                                numero, letra = split_numero(val)
                                numero_list.append(numero)
                                letra_list.append(letra)

                            df_xls['Numero'] = numero_list
                            df_xls.insert(num_pos + 1, 'Letra', letra_list)

                        # Formatear CUIT con guiones (XX-XXXXXXXX-X)
                        if 'C.U.I.T.' in df_xls.columns:
                            def fmt_cuit(val):
                                s = str(val).replace('-', '').replace('.', '').replace(' ', '').strip()
                                # Quitar .0 si viene de float
                                if s.endswith('.0'):
                                    s = s[:-2]
                                if len(s) == 11 and s.isdigit():
                                    return f"{s[:2]}-{s[2:10]}-{s[10]}"
                                return s
                            df_xls['C.U.I.T.'] = df_xls['C.U.I.T.'].apply(fmt_cuit)

                        # Rellenar y convertir columnas monetarias a numérico
                        for col_fill in ['Neto', 'Iva', 'Sobretasa', 'Retenciones']:
                            if col_fill in df_xls.columns:
                                df_xls[col_fill] = pd.to_numeric(df_xls[col_fill], errors='coerce').fillna(0)

                    total_valid = mask_header.sum()
                    not_found = total_valid - matched

                    st.success("✓  Cruce completado")

                    # Stats
                    st.markdown(f"""
                    <div class="stats-row">
                        <div class="stat-chip">
                            <span class="stat-val">{total_valid}</span>
                            <span class="stat-lbl">Comprobantes</span>
                        </div>
                        <div class="stat-chip">
                            <span class="stat-val">{matched}</span>
                            <span class="stat-lbl">Matcheados</span>
                        </div>
                        <div class="stat-chip">
                            <span class="stat-val">{not_found}</span>
                            <span class="stat-lbl">No encontrados</span>
                        </div>
                    </div>
                    """, unsafe_allow_html=True)

                    if not_found > 0:
                        st.warning(f"**{not_found}** comprobantes del Excel no fueron encontrados en el TXT")

                    # ── Generar Excel de salida con formato ──────────────
                    from openpyxl.styles import Border, Side, Font, PatternFill, Alignment
                    from openpyxl.utils import get_column_letter

                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        # startrow=4 → encabezados columna en fila 5, datos desde fila 6
                        df_xls.to_excel(writer, sheet_name='Movimientos', index=False, startrow=4)
                        ws = writer.sheets['Movimientos']

                        total_cols = len(df_xls.columns)
                        last_col_letter = get_column_letter(total_cols)
                        center_align = Alignment(horizontal='center', vertical='center')

                        # ── Encabezado con datos del cliente ──────────────
                        ws.merge_cells(f'A1:{last_col_letter}1')
                        ws['A1'] = meta_txt.get('razon_social', 'CONTRIBUYENTE').upper()
                        ws['A1'].font = Font(bold=True, size=14, color='FFFFFF')
                        ws['A1'].fill = PatternFill('solid', fgColor='2F5496')
                        ws['A1'].alignment = center_align

                        ws.merge_cells(f'A2:{last_col_letter}2')
                        tipo_rep = meta_txt.get('tipo_reporte', 'COMPRAS')
                        ws['A2'] = tipo_rep.upper()
                        ws['A2'].font = Font(bold=True, size=12, color='C00000')
                        ws['A2'].alignment = center_align

                        ws.merge_cells(f'A3:{last_col_letter}3')
                        ws['A3'] = f"CUIT: {meta_txt.get('cuit_empresa', '')} | Periodo: {meta_txt.get('periodo', '')}"
                        ws['A3'].font = Font(bold=True, size=11, color='2F5496')
                        ws['A3'].alignment = center_align

                        # ── Estilo encabezados de columna (fila 5) ────────
                        header_font = Font(bold=True, size=10, color='FFFFFF')
                        header_fill = PatternFill('solid', fgColor='4472C4')
                        header_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
                        for col_idx in range(1, total_cols + 1):
                            cell = ws.cell(row=5, column=col_idx)
                            cell.font = header_font
                            cell.fill = header_fill
                            cell.alignment = header_align

                        # ── Formato numérico con 2 decimales, rojo si negativo ──
                        num_fmt_red = '$#,##0.00;[Red]-$#,##0.00'
                        col_list_xls = list(df_xls.columns)
                        money_cols_xls = ['Neto', 'Iva', 'Sobretasa', 'Retenciones', 'Total']
                        money_indices = [col_list_xls.index(c) + 1 for c in money_cols_xls if c in col_list_xls]

                        data_start_row = 6
                        for row in range(data_start_row, len(df_xls) + data_start_row):
                            for col_idx in money_indices:
                                cell = ws.cell(row=row, column=col_idx)
                                cell.number_format = num_fmt_red

                    output.seek(0)

                    xls_name = Path(uploaded_xls_concepto.name).stem
                    st.download_button(
                        label="↓  Descargar Movimientos",
                        data=output,
                        file_name=f"Movimientos.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )

            except Exception as e:
                st.error(f"Error al procesar: {str(e)}")
                st.exception(e)

        st.markdown('</div>', unsafe_allow_html=True)

    else:
        st.markdown("""
        <div style="
            text-align: center;
            padding: 2rem 1rem;
            font-family: 'Space Mono', monospace;
            font-size: 0.72rem;
            color: #6b7280;
            letter-spacing: 0.12em;
        ">
            SUBÍ AMBOS ARCHIVOS · TXT + EXCEL SISTEMA
        </div>
        """, unsafe_allow_html=True)
