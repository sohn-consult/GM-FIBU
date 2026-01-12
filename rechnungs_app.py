import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO, StringIO
from datetime import datetime
import numpy as np

# --- 1. KONFIGURATION & DESIGN ---
st.set_page_config(page_title="Sohn-Consult | BI Dashboard", page_icon="👔", layout="wide")

st.markdown("""
    <style>
    .stApp { background-color: #FFFFFF; }
    [data-testid="stSidebar"] { background-color: #F8FAFC; border-right: 1px solid #E2E8F0; }
    h1, h2, h3 { color: #1E3A8A; font-family: 'Inter', sans-serif; }
    .stMetric { background-color: #FFFFFF; padding: 20px; border-radius: 10px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); border-top: 5px solid #1E3A8A; }
    .stButton>button { background-color: #1E3A8A; color: white; border-radius: 8px; font-weight: bold; width: 100%; height: 3em; }
    </style>
    """, unsafe_allow_html=True)

# Hilfsfunktion für Euro-Formatierung
def format_euro(val):
    if pd.isna(val) or val == "-": return "0,00 €"
    return f"{val:,.2f} €".replace(",", "X").replace(".", ",").replace("X", ".")

st.title("👔 Sohn-Consult | Strategische Fibu-Analyse")
st.markdown("---")

# --- 2. DATEI-UPLOAD ---
uploaded_file = st.file_uploader("📂 Excel oder DATEV-CSV hochladen", type=["xlsx", "csv"])

if uploaded_file:
    with st.sidebar:
        st.markdown("### 📥 1. Daten-Quelle")
        mode = st.radio("Format", ["Standard Excel/CSV", "DATEV-Export (CSV)"])
        
        try:
            # Datei einlesen
            if mode == "DATEV-Export (CSV)":
                raw_bytes = uploaded_file.getvalue()
                content = raw_bytes.decode('latin-1', errors='ignore')
                df = pd.read_csv(StringIO(content), sep=None, engine='python', skiprows=1)
                if df.empty or "Umsatz" not in df.columns:
                    df = pd.read_csv(StringIO(content), sep=None, engine='python')
            else:
                if uploaded_file.name.endswith('.csv'):
                    df = pd.read_csv(uploaded_file, sep=None, engine='python')
                else:
                    header_row = st.number_input("Header-Zeile (Spaltennamen)", min_value=1, value=3)
                    xl = pd.ExcelFile(uploaded_file)
                    sheet = st.selectbox("Blatt", xl.sheet_names)
                    df = pd.read_excel(uploaded_file, sheet_name=sheet, header=header_row-1)

            # Reinigung
            df = df.loc[:, ~df.columns.str.contains('^Unnamed')].dropna(how='all', axis=0)
            
            if df.empty:
                st.error("Die Tabelle scheint leer zu sein. Bitte prüfe die Header-Zeile.")
                st.stop()

            all_cols = df.columns.tolist()

            st.markdown("### 📍 2. Zuordnung")
            def get_idx(keys, default=0):
                for i, col in enumerate(all_cols):
                    if any(k.lower() in str(col).lower() for k in keys): return i
                return default if default < len(all_cols) else 0

            c_datum = st.selectbox("Datum", all_cols, index=get_idx(["datum", "belegdat"]))
            c_faellig = st.selectbox("Fälligkeit", all_cols, index=get_idx(["fällig", "termin"]))
            c_nr = st.selectbox("RE-Nummer", all_cols, index=get_idx(["belegfeld", "nummer", "re-nr"]))
            c_kunde = st.selectbox("Kunde", all_cols, index=get_idx(["name", "kunde"]))
            c_betrag = st.selectbox("Betrag", all_cols, index=get_idx(["brutto", "umsatz", "betrag"]))
            c_bezahlt = st.selectbox("Zahldatum", all_cols, index=get_idx(["bezahlt", "eingang"]))

            # Transformation (Sicher)
            df[c_datum] = pd.to_datetime(df[c_datum], errors='coerce')
            df[c_faellig] = pd.to_datetime(df[c_faellig], errors='coerce')
            df[c_bezahlt] = pd.to_datetime(df[c_bezahlt], errors='coerce')
            
            if df[c_betrag].dtype == 'object':
                df[c_betrag] = pd.to_numeric(df[c_betrag].astype(str).str.replace('.', '').str.replace(',', '.'), errors='coerce')
            
            # Daten filtern
            df = df.dropna(subset=[c_datum, c_betrag])
            
            if df.empty:
                st.warning("Keine gültigen Zeilen nach der Datum/Betrag-Prüfung gefunden.")
                st.stop()

            st.markdown("### 🔍 3. Filter")
            kunden = sorted(df[c_kunde].dropna().unique().tolist())
            sel_kunden = st.multiselect("Kunden", options=kunden, default=kunden)
            
            min_d = df[c_datum].min().date()
            max_d = df[c_datum].max().date()
            date_range = st.date_input("Zeitraum", [min_d, max_d])

            st.markdown("---")
            start_btn = st.button("🚀 ANALYSE STARTEN")

        except Exception as e:
            st.error(f"Kritischer Fehler beim Laden: {e}")
            st.stop()

    # --- 3. HAUPTBEREICH ---
    if start_btn and len(date_range) == 2:
        mask = (df[c_datum].dt.date >= date_range[0]) & (df[c_datum].dt.date <= date_range[1]) & (df[c_kunde].isin(sel_kunden))
        f_df = df[mask].copy()

        if f_df.empty:
            st.warning("Keine Daten für diesen Filter gefunden.")
        else:
            t1, t2, t3, t4 = st.tabs(["📊 Performance", "🔴 Mahnwesen", "💎 ABC-Check", "🔍 Forensik"])

            with t1:
                col1, col2, col3 = st.columns(3)
                col1.metric("Gesamtumsatz", format_euro(f_df[c_betrag].sum()))
                col2.metric("Offene Posten", format_euro(f_df[f_df[c_bezahlt].isna()][c_betrag].sum()))
                col3.metric("Anzahl Belege", len(f_df))
                
                f_df['Monat'] = f_df[c_datum].dt.strftime('%Y-%m')
                chart = f_df.groupby('Monat')[c_betrag].sum().reset_index()
                st.plotly_chart(px.bar(chart, x='Monat', y=c_betrag, color_discrete_sequence=['#1E3A8A']), use_container_width=True)

            with t2:
                offen = f_df[f_df[c_bezahlt].isna()].copy()
                if not offen.empty:
                    offen['Verzug'] = (pd.Timestamp(datetime.now().date()) - offen[c_faellig]).dt.days
                    st.dataframe(offen[[c_datum, c_kunde, c_betrag, 'Verzug']].sort_values(by='Verzug', ascending=False),
                                 column_config={c_betrag: st.column_config.NumberColumn(format="%.2f €")}, use_container_width=True)
                else: st.success("Alles bezahlt.")

            with t3:
                abc = f_df.groupby(c_kunde)[c_betrag].sum().reset_index().sort_values(by=c_betrag, ascending=False)
                st.plotly_chart(px.pie(abc, values=c_betrag, names=c_kunde, hole=0.4), use_container_width=True)

            with t4:
                st.subheader("Integritätsprüfung")
                # Nummernkreis-Check
                try:
                    nums = pd.to_numeric(f_df[c_nr], errors='coerce').dropna().astype(int).sort_values().unique()
                    if len(nums) > 1:
                        diffs = np.diff(nums)
                        if any(diffs > 1): st.warning("⚠️ Lücken im Nummernkreis gefunden.")
                        else: st.success("✅ RE-Nummern lückenlos.")
                except: st.info("Check nicht möglich.")
                
                # Logik-Check
                err = f_df[f_df[c_bezahlt] < f_df[c_datum]]
                if not err.empty: st.error("Logik-Fehler: Zahlung vor Rechnung!"); st.dataframe(err)
                else: st.success("Datum-Logik ok.")

    else:
        st.info("Oben Datei laden und links Filter einstellen. Dann 'Analyse starten'.")
        st.markdown("### 📄 Vorschau der Rohdaten (erste 5 Zeilen):")
        st.write(df.head(5)) # Hilft die richtige Header-Zeile zu finden
