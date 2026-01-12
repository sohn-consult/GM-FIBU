import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO, StringIO
from datetime import datetime
import numpy as np

# --- 1. DESIGN & KONFIGURATION ---
st.set_page_config(page_title="Sohn-Consult | BI Dashboard", page_icon="👔", layout="wide")

# CSS für maximale Lesbarkeit und Sohn-Consult Branding
st.markdown("""
    <style>
    .stApp { background-color: #FFFFFF; }
    [data-testid="stSidebar"] { background-color: #F1F5F9; border-right: 1px solid #CBD5E1; }
    [data-testid="stSidebar"] .stMarkdown p, [data-testid="stSidebar"] label { 
        color: #0F172A !important; 
        font-weight: 600 !important; 
    }
    h1, h2, h3 { color: #1E3A8A; font-family: 'Inter', sans-serif; font-weight: 700; }
    .stMetric { 
        background-color: #FFFFFF; padding: 20px; border-radius: 10px; 
        box-shadow: 0 4px 6px rgba(0,0,0,0.05); border-top: 5px solid #1E3A8A; 
    }
    .stButton>button { 
        background-color: #1E3A8A; color: white; border-radius: 8px; 
        font-weight: bold; width: 100%; height: 3.5em; 
    }
    .stTabs [aria-selected="true"] { background-color: #1E3A8A !important; color: white !important; }
    </style>
    """, unsafe_allow_html=True)

# Hilfsfunktion für Euro-Formatierung
def format_euro(val):
    if pd.isna(val): return "0,00 €"
    return f"{val:,.2f} €".replace(",", "X").replace(".", ",").replace("X", ".")

st.title("👔 Sohn-Consult | Business Intelligence")
st.caption("Professionelle Fibu-Analyse & Strategische Auswertung")
st.markdown("---")

# --- 2. DATEI-UPLOAD ---
uploaded_file = st.file_uploader("📂 Fibu-Daten hochladen (Excel oder DATEV-CSV)", type=["xlsx", "csv"])

if uploaded_file:
    with st.sidebar:
        st.markdown("### 📥 1. Daten-Quelle")
        mode = st.radio("Format auswählen", ["Standard Excel/CSV", "DATEV-Export (CSV)"])
        
        try:
            # Datei einlesen
            if mode == "DATEV-Export (CSV)":
                raw_bytes = uploaded_file.getvalue()
                content = raw_bytes.decode('latin-1', errors='ignore')
                df_raw = pd.read_csv(StringIO(content), sep=None, engine='python', skiprows=1)
                if df_raw.empty or "Umsatz" not in df_raw.columns:
                    df_raw = pd.read_csv(StringIO(content), sep=None, engine='python')
            else:
                if uploaded_file.name.endswith('.csv'):
                    df_raw = pd.read_csv(uploaded_file, sep=None, engine='python')
                else:
                    header_row = st.number_input("Header-Zeile (Spaltennamen)", min_value=1, value=3)
                    xl = pd.ExcelFile(uploaded_file)
                    sheet = st.selectbox("Tabellenblatt", xl.sheet_names)
                    df_raw = pd.read_excel(uploaded_file, sheet_name=sheet, header=header_row-1)

            # Basis-Reinigung
            df_raw = df_raw.loc[:, ~df_raw.columns.str.contains('^Unnamed')].dropna(how='all', axis=0)
            all_cols = df_raw.columns.tolist()

            st.markdown("### 📍 2. Spalten-Zuordnung")
            def find_idx(keys, default=0):
                for i, c in enumerate(all_cols):
                    if any(k.lower() in str(c).lower() for k in keys): return i
                return default if default < len(all_cols) else 0

            # Zuordnungen
            c_datum = st.selectbox("Rechnungsdatum", all_cols, index=find_idx(["datum", "belegdat"]))
            c_faellig = st.selectbox("Fälligkeitsdatum", all_cols, index=find_idx(["fällig", "termin"]))
            c_nr = st.selectbox("RE-Nummer", all_cols, index=find_idx(["belegfeld", "nummer", "re-nr"]))
            c_kunde = st.selectbox("Kunde", all_cols, index=find_idx(["name", "kunde"]))
            c_betrag = st.selectbox("Betrag (Brutto)", all_cols, index=find_idx(["brutto", "umsatz", "betrag"]))
            c_bezahlt = st.selectbox("Zahldatum", all_cols, index=find_idx(["bezahlt", "eingang"]))

            # --- VORBEREITUNG DER GEFILTERTEN DATEN ---
            df_work = df_raw.copy()
            df_work[c_datum] = pd.to_datetime(df_work[c_datum], errors='coerce')
            df_work[c_faellig] = pd.to_datetime(df_work[c_faellig], errors='coerce')
            df_work[c_bezahlt] = pd.to_datetime(df_work[c_bezahlt], errors='coerce')
            
            if df_work[c_betrag].dtype == 'object':
                df_work[c_betrag] = pd.to_numeric(df_work[c_betrag].astype(str).str.replace('.', '').str.replace(',', '.'), errors='coerce')
            
            # Wichtig: Zeilen ohne Datum/Betrag entfernen, bevor Filter gebaut wird
            df_work = df_work.dropna(subset=[c_datum, c_betrag])

            st.markdown("### 🔍 3. Filter")
            if not df_work.empty:
                # Kundenliste aus der gewählten Spalte extrahieren
                kunden_options = sorted(df_work[c_kunde].dropna().unique().tolist())
                
                # DER FIX: Multiselect mit Fallback
                sel_kunden = st.multiselect(
                    "Kunden auswählen", 
                    options=kunden_options, 
                    default=kunden_options
                )
                
                # Zeitbereich
                min_d, max_d = df_work[c_datum].min().date(), df_work[c_datum].max().date()
                date_range = st.date_input("Zeitraum", [min_d, max_d])

                st.markdown("---")
                # Start-Knopf
                start_btn = st.button("🚀 ANALYSE STARTEN", use_container_width=True)
            else:
                st.warning("Keine Daten gefunden. Prüfe Header & Spalten.")
                start_btn = False

        except Exception as e:
            st.error(f"Fehler: {e}")
            start_btn = False

    # --- 3. HAUPTBEREICH: ANALYSE ---
    if start_btn and len(date_range) == 2:
        # Finale Filterung
        mask = (df_work[c_datum].dt.date >= date_range[0]) & \
               (df_work[c_datum].dt.date <= date_range[1]) & \
               (df_work[c_kunde].isin(sel_kunden))
        f_df = df_work[mask].copy()

        if f_df.empty:
            st.warning("Keine Daten für die gewählten Filter vorhanden.")
        else:
            tabs = st.tabs(["📊 Performance", "🔴 Mahnwesen", "💎 ABC/Risk", "🔍 Forensik"])

            with tabs[0]:
                m1, m2, m3 = st.columns(3)
                offen_mask = f_df[c_bezahlt].isna()
                m1.metric("Gesamtumsatz", format_euro(f_df[c_betrag].sum()))
                m2.metric("Offene Posten", format_euro(f_df[offen_mask][c_betrag].sum()))
                m3.metric("Belege", len(f_df))
                
                f_df['Monat'] = f_df[c_datum].dt.strftime('%Y-%m')
                chart = f_df.groupby('Monat')[c_betrag].sum().reset_index()
                st.plotly_chart(px.bar(chart, x='Monat', y=c_betrag, color_discrete_sequence=['#1E3A8A']), use_container_width=True)

            with tabs[1]:
                offen = f_df[offen_mask].copy()
                if not offen.empty:
                    offen['Verzug'] = (pd.Timestamp(datetime.now().date()) - offen[c_faellig]).dt.days
                    st.dataframe(offen[[c_datum, c_kunde, c_betrag, 'Verzug']].sort_values(by='Verzug', ascending=False),
                                 column_config={c_betrag: st.column_config.NumberColumn(format="%.2f €")}, use_container_width=True)
                else: st.success("Alles bezahlt.")

            with tabs[2]:
                abc = f_df.groupby(c_kunde)[c_betrag].sum().reset_index().sort_values(by=c_betrag, ascending=False)
                st.plotly_chart(px.pie(abc, values=c_betrag, names=c_kunde, hole=0.4), use_container_width=True)
                top_3 = (abc[c_betrag].head(3).sum() / abc[c_betrag].sum()) * 100
                st.metric("Klumpenrisiko (Top 3)", f"{top_3:.1f}%")

            with tabs[3]:
                st.subheader("Forensik & Logik")
                # Nummernkreis
                try:
                    nums = pd.to_numeric(f_df[c_nr], errors='coerce').dropna().astype(int).sort_values().unique()
                    if len(nums) > 1:
                        missing = np.setdiff1d(np.arange(nums.min(), nums.max() + 1), nums)
                        if len(missing) > 0: st.warning(f"⚠️ {len(missing)} Rechnungsnummern fehlen.")
                        else: st.success("✅ RE-Nummern lückenlos.")
                except: st.info("Check nicht verfügbar.")
                
                # Logik
                err = f_df[f_df[c_bezahlt] < f_df[c_datum]]
                if not err.empty: st.error("❌ Zahlung vor Rechnung gefunden!"); st.dataframe(err)

    else:
        st.info("Bitte Datei laden und links Filter einstellen. Dann 'Analyse starten'.")
        if 'df_raw' in locals():
            st.markdown("### 📄 Vorschau der Daten:")
            st.dataframe(df_raw.head(5))
