import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO, StringIO
from datetime import datetime
import numpy as np

# --- 1. DESIGN & KONFIGURATION (Update 2026) ---
st.set_page_config(page_title="Sohn-Consult | Strategic BI", page_icon="👔", layout="wide")

st.markdown("""
    <style>
    .stApp { background-color: #FFFFFF; }
    [data-testid="stSidebar"] { background-color: #F1F5F9; border-right: 1px solid #CBD5E1; }
    [data-testid="stSidebar"] .stMarkdown p, [data-testid="stSidebar"] label { color: #0F172A !important; font-weight: 600 !important; }
    h1, h2, h3 { color: #1E3A8A; font-family: 'Inter', sans-serif; font-weight: 700; }
    .stMetric { background-color: #FFFFFF; padding: 20px; border-radius: 10px; box-shadow: 0 4px 6px rgba(0,0,0,0.05); border-top: 5px solid #1E3A8A; }
    .stButton>button { background-color: #1E3A8A; color: white; border-radius: 8px; font-weight: bold; width: 100%; height: 3.5em; }
    .stTabs [aria-selected="true"] { background-color: #1E3A8A !important; color: white !important; }
    </style>
    """, unsafe_allow_html=True)

def format_euro(val):
    if pd.isna(val) or val is None: return "0,00 €"
    return f"{val:,.2f} €".replace(",", "X").replace(".", ",").replace("X", ".")

def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Sohn_Consult_Analyse')
    return output.getvalue()

st.title("👔 Sohn-Consult | Strategic BI")
st.caption("Fibu-Analyse, Bank-Abgleich & Forensik (Stabilisierte Version 2026.2)")
st.markdown("---")

# --- 2. MULTI-UPLOAD BEREICH ---
col_u1, col_u2 = st.columns(2)
with col_u1:
    fibu_file = st.file_uploader("📂 1. Fibu/Debitoren-Datei laden", type=["xlsx", "csv"])
with col_u2:
    bank_file = st.file_uploader("🏦 2. OPTIONAL: Bankumsätze (CSV) laden", type=["csv"])

if fibu_file:
    with st.sidebar:
        st.markdown("### 📥 Konfiguration")
        mode = st.radio("Dateiformat", ["Standard Excel/CSV", "DATEV-Export (CSV)"])
        
        try:
            if mode == "DATEV-Export (CSV)":
                raw_bytes = fibu_file.getvalue()
                df_raw = pd.read_csv(StringIO(raw_bytes.decode('latin-1', errors='ignore')), sep=None, engine='python', skiprows=1)
            else:
                if fibu_file.name.endswith('.csv'):
                    df_raw = pd.read_csv(fibu_file, sep=None, engine='python')
                else:
                    header_row = st.number_input("Header-Zeile", min_value=1, value=3)
                    df_raw = pd.read_excel(fibu_file, sheet_name=0, header=header_row-1)

            # --- SICHERER SPALTENFILTER (FIX FÜR LOG-FEHLER) ---
            # Wir stellen sicher, dass alle Spaltennamen Strings sind, bevor wir filtern
            df_raw.columns = [str(c) for c in df_raw.columns]
            valid_cols = [c for c in df_raw.columns if "Unnamed" not in c]
            df_raw = df_raw[valid_cols].dropna(how='all', axis=0)
            
            all_cols = df_raw.columns.tolist()

            st.markdown("### 📍 Spalten-Zuordnung")
            def find_idx(keys, default=0):
                for i, c in enumerate(all_cols):
                    if any(k.lower() in str(c).lower() for k in keys): return i
                return default

            c_datum = st.selectbox("Rechnungsdatum", all_cols, index=find_idx(["datum"]))
            c_faellig = st.selectbox("Fälligkeitsdatum", all_cols, index=find_idx(["fällig", "termin"]))
            c_nr = st.selectbox("RE-Nummer", all_cols, index=find_idx(["nummer", "belegfeld"]))
            c_kunde = st.selectbox("Kunde", all_cols, index=find_idx(["kunde", "name"]))
            c_betrag = st.selectbox("Betrag (Brutto)", all_cols, index=find_idx(["brutto", "betrag", "umsatz"]))
            c_bezahlt = st.selectbox("Zahldatum (leer = offen)", all_cols, index=find_idx(["gezahlt", "bezahlt", "eingang"]))

            # --- DATEN-TRANSFORMATION ---
            df_work = df_raw.copy()
            if df_work[c_betrag].dtype == 'object':
                df_work[c_betrag] = pd.to_numeric(df_work[c_betrag].astype(str).str.replace('.', '').str.replace(',', '.'), errors='coerce')
            
            df_work['Fällig_Display'] = df_work[c_faellig].astype(str)
            for col in [c_datum, c_faellig, c_bezahlt]:
                df_work[col] = pd.to_datetime(df_work[col], errors='coerce')
            
            df_work = df_work.dropna(subset=[c_datum, c_betrag])

            # Zeitraum-Filter
            min_d, max_d = df_work[c_datum].min().date(), df_work[c_datum].max().date()
            date_range = st.date_input("Zeitraum wählen", [min_d, max_d])

            st.markdown("---")
            start_btn = st.button("🚀 ANALYSE STARTEN", use_container_width=True)

        except Exception as e:
            st.error(f"Fehler beim Einlesen: {e}")
            start_btn = False

    # --- 3. HAUPTBEREICH ---
    if start_btn and len(date_range) == 2:
        mask = (df_work[c_datum].dt.date >= date_range[0]) & (df_work[c_datum].dt.date <= date_range[1])
        f_df = df_work[mask].copy()

        if f_df.empty:
            st.warning("Keine Daten im gewählten Zeitraum.")
        else:
            # Offene Posten Logik
            offen_mask = f_df[c_bezahlt].isna()
            offene_df = f_df[offen_mask].copy()
            
            tabs = st.tabs(["📊 Performance", "🔴 Offene Posten", "🔍 Forensik", "🏦 Bank-Abgleich"])

            with tabs[0]:
                m1, m2, m3 = st.columns(3)
                total_rev = f_df[c_betrag].sum()
                open_rev = offene_df[c_betrag].sum()
                m1.metric("Umsatz gesamt", format_euro(total_rev))
                m2.metric("Offene Forderungen", format_euro(open_rev))
                m3.metric("Anzahl Belege", len(f_df))
                
                f_df['Monat'] = f_df[c_datum].dt.strftime('%Y-%m')
                st.plotly_chart(px.bar(f_df.groupby('Monat')[c_betrag].sum().reset_index(), x='Monat', y=c_betrag, color_discrete_sequence=['#1E3A8A']), width='stretch')

            with tabs[1]:
                st.subheader("Übersicht unbezahlte Rechnungen")
                if not offene_df.empty:
                    today = pd.Timestamp(datetime.now().date())
                    offene_df['Verzug'] = (today - offene_df[c_faellig]).dt.days
                    disp = offene_df[[c_datum, 'Fällig_Display', c_kunde, c_betrag, 'Verzug']].copy()
                    st.dataframe(disp.sort_values(by='Verzug', ascending=False), 
                                 column_config={c_betrag: st.column_config.NumberColumn(format="%.2f €")}, width='stretch')
                else:
                    st.success("Alle Rechnungen sind bezahlt.")

            with tabs[2]:
                st.subheader("Strategische Analyse")
                c_abc, c_for = st.columns(2)
                with c_abc:
                    abc = f_df.groupby(c_kunde)[c_betrag].sum().reset_index().sort_values(by=c_betrag, ascending=False)
                    st.plotly_chart(px.pie(abc, values=c_betrag, names=c_kunde, hole=0.4, title="Kundenanteile"), width='stretch')
                with c_for:
                    st.markdown("**Forensik-Check**")
                    err = f_df[f_df[c_bezahlt] < f_df[c_datum]]
                    if not err.empty: st.error(f"❌ {len(err)} Zahlungen vor Rechnungsdatum gefunden!")
                    else: st.success("Datum-Logik einwandfrei.")

            with tabs[3]:
                st.subheader("Kontoauszugs-Abgleich")
                if bank_file:
                    df_bank = pd.read_csv(bank_file, sep=None, engine='python')
                    st.success("Bankdaten geladen.")
                    st.dataframe(df_bank.head(10), width='stretch')
                else:
                    st.info("Laden Sie oben eine Bank-CSV hoch, um Rechnungen abzugleichen.")
    else:
        st.info("Bitte Datei laden und Analyse starten.")
