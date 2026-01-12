import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO, StringIO
from datetime import datetime

# --- 1. DESIGN & KONFIGURATION ---
st.set_page_config(page_title="Sohn-Consult BI", page_icon="📊", layout="wide")

# Optimiertes CSS für maximale Lesbarkeit
st.markdown("""
    <style>
    /* Haupt-Hintergrund */
    .stApp {
        background-color: #FFFFFF;
    }
    /* Sidebar: Hell mit blauem Text für besseren Kontrast */
    [data-testid="stSidebar"] {
        background-color: #F0F4F8; 
        border-right: 1px solid #D1D5DB;
    }
    /* Alle Texte in der Sidebar auf Dunkelblau setzen */
    [data-testid="stSidebar"] .stMarkdown p, 
    [data-testid="stSidebar"] label, 
    [data-testid="stSidebar"] .stSelectbox label,
    [data-testid="stSidebar"] .stRadio label {
        color: #1E3A8A !important;
        font-weight: 600 !important;
    }
    /* Titel-Styling */
    h1 {
        color: #1E3A8A;
        font-family: 'Inter', sans-serif;
        font-weight: 800;
        letter-spacing: -1px;
    }
    /* Metric Cards: Weiß mit blauem Akzent-Balken oben */
    .stMetric {
        background-color: #FFFFFF;
        padding: 20px;
        border-radius: 8px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
        border-top: 4px solid #1E3A8A;
    }
    /* Button: Kräftiges Blau */
    .stButton>button {
        background-color: #1E3A8A;
        color: white;
        border-radius: 6px;
        border: none;
        padding: 12px;
        width: 100%;
        transition: all 0.3s;
    }
    .stButton>button:hover {
        background-color: #111827;
        box-shadow: 0 4px 12px rgba(0,0,0,0.2);
    }
    /* Tabs: Schlicht und modern */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
    }
    .stTabs [data-baseweb="tab"] {
        background-color: #F3F4F6;
        border-radius: 4px 4px 0 0;
        padding: 10px 20px;
        color: #4B5563;
    }
    .stTabs [aria-selected="true"] {
        background-color: #1E3A8A !important;
        color: white !important;
    }
    </style>
    """, unsafe_allow_html=True)

# Header
st.title("📊 Sohn-Consult | Business Intelligence")
st.markdown("---")

def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Sohn_Consult_Fibu')
    return output.getvalue()

# --- 2. DATEI-UPLOAD ---
uploaded_file = st.file_uploader("📂 Fibu-Daten hochladen (Excel oder DATEV-CSV)", type=["xlsx", "csv"])

if uploaded_file:
    with st.sidebar:
        st.markdown("### 📥 1. Daten-Quelle")
        mode = st.radio("Format auswählen", ["Excel / Standard CSV", "DATEV-Export (CSV)"])
        
        try:
            if mode == "DATEV-Export (CSV)":
                raw_bytes = uploaded_file.getvalue()
                content = raw_bytes.decode('latin-1')
                df = pd.read_csv(StringIO(content), sep=None, engine='python', skiprows=1)
                if "Umsatz" not in df.columns:
                    df = pd.read_csv(StringIO(content), sep=None, engine='python')
            else:
                if uploaded_file.name.endswith('.csv'):
                    df = pd.read_csv(uploaded_file, sep=None, engine='python')
                else:
                    header_row = st.number_input("Header-Zeile (Spaltennamen)", min_value=1, value=3)
                    xl = pd.ExcelFile(uploaded_file)
                    sheet = st.selectbox("Tabellenblatt", xl.sheet_names)
                    df = pd.read_excel(uploaded_file, sheet_name=sheet, header=header_row-1)

            df = df.loc[:, ~df.columns.str.contains('^Unnamed')].dropna(how='all', axis=0)
            all_cols = df.columns.tolist()

            st.markdown("### 📍 2. Zuordnung")
            def find_col(keys, default):
                for i, col in enumerate(all_cols):
                    if any(k.lower() in str(col).lower() for k in keys): return i
                return default if default < len(all_cols) else 0

            c_datum = st.selectbox("Rechnungsdatum", all_cols, index=find_col(["datum", "belegdat"], 2))
            c_faellig = st.selectbox("Fälligkeitsdatum", all_cols, index=find_col(["fällig", "termin"], 3))
            c_nr = st.selectbox("RE-Nummer", all_cols, index=find_col(["belegfeld", "nummer", "re-nr"], 1))
            c_kunde = st.selectbox("Kunde", all_cols, index=find_col(["name", "kunde", "gegenkonto"], 0))
            c_betrag = st.selectbox("Betrag (Brutto)", all_cols, index=find_col(["brutto", "umsatz", "betrag"], 16))
            c_bezahlt = st.selectbox("Zahldatum (leer=offen)", all_cols, index=find_col(["bezahlt", "eingang", "ausgleich"], 17))

            # Transformation
            for col in [c_datum, c_faellig, c_bezahlt]:
                df[col] = pd.to_datetime(df[col], errors='coerce')
            if df[c_betrag].dtype == 'object':
                df[c_betrag] = pd.to_numeric(df[c_betrag].astype(str).str.replace('.', '').str.replace(',', '.'), errors='coerce')
            df = df.dropna(subset=[c_datum, c_betrag])

            st.markdown("### 🔍 3. Filter")
            kunden = sorted(df[c_kunde].dropna().unique().tolist())
            sel_kunden = st.multiselect("Kunden auswählen", options=kunden, default=kunden)
            date_range = st.date_input("Zeitraum", [df[c_datum].min().date(), df[c_datum].max().date()])

            st.markdown("---")
            start_btn = st.button("🚀 ANALYSE STARTEN")

        except Exception as e:
            st.error(f"Fehler: {e}")
            start_btn = False

    # --- 3. ANALYSE ---
    if start_btn and len(date_range) == 2:
        mask = (df[c_datum].dt.date >= date_range[0]) & (df[c_datum].dt.date <= date_range[1]) & (df[c_kunde].isin(sel_kunden))
        f_df = df[mask].copy()

        t1, t2, t3, t4, t5 = st.tabs(["📊 Performance", "🔴 Mahnwesen", "📅 Cashflow", "💎 ABC-Analyse", "🔍 Forensik"])

        with t1:
            m1, m2, m3 = st.columns(3)
            offen = f_df[f_df[c_bezahlt].isna()]
            m1.metric("Umsatz", f"{f_df[c_betrag].sum():,.2f} €")
            m2.metric("Offen", f"{offen[c_betrag].sum():,.2f} €")
            m3.metric("Belege", len(f_df))
            
            f_df['Monat'] = f_df[c_datum].dt.strftime('%Y-%m')
            st.plotly_chart(px.bar(f_df.groupby('Monat')[c_betrag].sum().reset_index(), x='Monat', y=c_betrag, color_discrete_sequence=['#1E3A8A']), use_container_width=True)

        with t2:
            st.subheader("Fälligkeiten")
            offen = offen.copy()
            offen['Verzug'] = (pd.Timestamp(datetime.now().date()) - offen[c_faellig]).dt.days
            st.dataframe(offen[[c_datum, c_faellig, c_kunde, c_betrag, 'Verzug']].sort_values(by='Verzug', ascending=False), use_container_width=True)

        with t3:
            cf = offen.groupby(c_faellig)[c_betrag].sum().reset_index()
            st.plotly_chart(px.line(cf, x=c_faellig, y=c_betrag, markers=True, title="Liquiditätsplanung", color_discrete_sequence=['#1E3A8A']), use_container_width=True)

        with t4:
            abc = f_df.groupby(c_kunde)[c_betrag].sum().reset_index().sort_values(by=c_betrag, ascending=False)
            st.plotly_chart(px.pie(abc, values=c_betrag, names=c_kunde, title="Umsatzverteilung"), use_container_width=True)

        with t5:
            st.subheader("Forensik")
            err = f_df[f_df[c_bezahlt] < f_df[c_datum]]
            if not err.empty: st.error("Zahlung vor Rechnungsdatum gefunden!"); st.dataframe(err)
            else: st.success("Daten-Logik ok.")

    else:
        st.info("Bitte Datei hochladen und 'Analyse starten' klicken.")
