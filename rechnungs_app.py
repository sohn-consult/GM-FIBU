import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO, StringIO
from datetime import datetime
import numpy as np

# --- 1. KONFIGURATION & BRANDING ---
st.set_page_config(page_title="Sohn-Consult | Strategic BI", page_icon="📈", layout="wide")

st.markdown("""
    <style>
    .stApp { background-color: #FFFFFF; }
    [data-testid="stSidebar"] { background-color: #F8FAFC; border-right: 1px solid #E2E8F0; }
    h1, h2, h3 { color: #1E3A8A; font-family: 'Inter', sans-serif; font-weight: 700; }
    .stMetric { background-color: #FFFFFF; padding: 20px; border-radius: 12px; border-top: 5px solid #1E3A8A; box-shadow: 0 4px 10px rgba(0,0,0,0.05); }
    .stButton>button { background-color: #1E3A8A; color: white; border-radius: 8px; font-weight: bold; height: 3.5em; width: 100%; }
    .stTabs [aria-selected="true"] { background-color: #1E3A8A !important; color: white !important; border-radius: 6px; }
    </style>
    """, unsafe_allow_html=True)

def format_euro(val):
    if pd.isna(val): return "0,00 €"
    return f"{val:,.2f} €".replace(",", "X").replace(".", ",").replace("X", ".")

st.title("👔 Sohn-Consult | Executive BI Dashboard")
st.caption("Universal-Tool für Fibu-Analyse, Bank-Abgleich & Strategie-Consulting")
st.markdown("---")

# --- 2. MULTI-UPLOAD BEREICH ---
col_u1, col_u2 = st.columns(2)
with col_u1:
    fibu_file = st.file_uploader("📂 1. Fibu/Debitoren-Datei laden", type=["xlsx", "csv"])
with col_u2:
    bank_file = st.file_uploader("🏦 2. OPTIONAL: Bankumsätze (CSV) laden", type=["csv"])

if fibu_file:
    with st.sidebar:
        st.markdown("### ⚙️ Konfiguration")
        header_idx = st.number_input("Header-Zeile Fibu", min_value=1, value=3)
        
        # Einlesen Fibu
        if fibu_file.name.endswith('.csv'):
            df_fibu = pd.read_csv(fibu_file, sep=None, engine='python')
        else:
            xl = pd.ExcelFile(fibu_file)
            sheet = st.selectbox("Blatt wählen", xl.sheet_names)
            df_fibu = pd.read_excel(fibu_file, sheet_name=sheet, header=header_idx-1)
        
        df_fibu = df_fibu.loc[:, ~df_fibu.columns.str.contains('^Unnamed')].dropna(how='all', axis=0)
        cols = df_fibu.columns.tolist()

        st.markdown("### 📍 Mapping")
        c_datum = st.selectbox("Belegdatum", cols, index=0)
        c_nr = st.selectbox("RE-Nummer", cols, index=1)
        c_kunde = st.selectbox("Kunde", cols, index=2)
        c_betrag = st.selectbox("Betrag (Brutto)", cols, index=3)
        
        # Daten-Formatierung
        df_fibu[c_datum] = pd.to_datetime(df_fibu[c_datum], errors='coerce')
        if df_fibu[c_betrag].dtype == 'object':
            df_fibu[c_betrag] = pd.to_numeric(df_fibu[c_betrag].astype(str).str.replace('.', '').str.replace(',', '.'), errors='coerce')
        df_fibu = df_fibu.dropna(subset=[c_datum, c_betrag])

        start_btn = st.button("🚀 ANALYSE STARTEN")

    if start_btn:
        t1, t2, t3, t4 = st.tabs(["📊 Performance", "🔍 Forensik & ABC", "🏦 Bank-Abgleich", "👔 Berater-Bericht"])

        # --- TAB 1: PERFORMANCE ---
        with t1:
            k1, k2, k3 = st.columns(3)
            total = df_fibu[c_betrag].sum()
            k1.metric("Gesamtumsatz", format_euro(total))
            k2.metric("Ø Rechnungswert", format_euro(df_fibu[c_betrag].mean()))
            k3.metric("Anzahl Belege", len(df_fibu))
            
            df_fibu['Monat'] = df_fibu[c_datum].dt.strftime('%Y-%m')
            st.plotly_chart(px.bar(df_fibu.groupby('Monat')[c_betrag].sum().reset_index(), x='Monat', y=c_betrag, color_discrete_sequence=['#1E3A8A']), use_container_width=True)

        # --- TAB 2: ABC & FORENSIK ---
        with t2:
            st.subheader("Klumpenrisiko-Check")
            abc = df_fibu.groupby(c_kunde)[c_betrag].sum().reset_index().sort_values(by=c_betrag, ascending=False)
            top_3 = (abc[c_betrag].head(3).sum() / total) * 100
            st.metric("Klumpenrisiko (Top 3)", f"{top_3:.1f}%")
            st.plotly_chart(px.pie(abc, values=c_betrag, names=c_kunde, hole=0.4), use_container_width=True)

        # --- TAB 3: BANK-ABGLEICH (ADD-ON) ---
        with t3:
            st.subheader("Automatischer Bank-Abgleich")
            if bank_file:
                df_bank = pd.read_csv(bank_file, sep=None, engine='python')
                st.write("Bankdaten erfolgreich geladen. Suche nach Übereinstimmungen...")
                
                # Beispielhafte Matching-Logik (Betragssuche)
                # Wir suchen in den Bankdaten nach Beträgen, die in der Fibu vorkommen
                bank_cols = df_bank.columns.tolist()
                b_betrag = st.selectbox("Spalte Bank-Betrag", bank_cols)
                b_text = st.selectbox("Spalte Verwendungszweck", bank_cols)
                
                # Konvertierung Bank-Betrag
                df_bank[b_betrag] = pd.to_numeric(df_bank[b_betrag].astype(str).str.replace(',', '.'), errors='coerce').abs()
                
                # Matching
                matched = df_fibu[df_fibu[c_betrag].isin(df_bank[b_betrag])]
                unmatched = df_fibu[~df_fibu[c_betrag].isin(df_bank[b_betrag])]
                
                m1, m2 = st.columns(2)
                m1.success(f"Gefundene Zahlungen: {len(matched)}")
                m2.error(f"Nicht auf Bank gefunden: {len(unmatched)}")
                
                st.write("**Nicht zugeordnete Rechnungen:**")
                st.dataframe(unmatched[[c_datum, c_kunde, c_betrag]], use_container_width=True)
            else:
                st.info("Laden Sie eine Bank-CSV hoch, um den Abgleich zu starten.")

        # --- TAB 4: BERATER-BERICHT ---
        with t4:
            st.subheader("Strategische Erkenntnisse")
            if total > 0:
                st.markdown(f"""
                ### 📝 Sohn-Consult Analyse-Zusammenfassung:
                1. **Umsatzstabilität:** Die Run-Rate basiert auf {len(df_fibu)} Belegen.
                2. **Risikoprofil:** Das Klumpenrisiko von {top_3:.1f}% deutet auf eine {'kritische' if top_3 > 50 else 'gesunde'} Abhängigkeit hin.
                3. **Liquiditäts-Tipp:** Prüfen Sie die Differenz im Bank-Abgleich auf 'Hidden Cashflows'.
                """)
                st.button("📄 Bericht als PDF exportieren (Simulation)")
else:
    st.info("👋 Willkommen! Bitte laden Sie zuerst die Fibu-Daten hoch.")
