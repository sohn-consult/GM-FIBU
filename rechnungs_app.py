import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO, StringIO
from datetime import datetime
import numpy as np

# --- 1. DESIGN & KONFIGURATION (Update 2026) ---
st.set_page_config(page_title="Sohn-Consult | BI Dashboard", page_icon="👔", layout="wide")

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
    if pd.isna(val): return "0,00 €"
    return f"{val:,.2f} €".replace(",", "X").replace(".", ",").replace("X", ".")

def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Sohn_Consult_Analyse')
    return output.getvalue()

st.title("👔 Sohn-Consult | Business Intelligence")
st.caption("Fibu-Analyse & Forensic Dashboard - Version 2026.1")
st.markdown("---")

# --- 2. DATEI-UPLOAD ---
uploaded_file = st.file_uploader("📂 Fibu-Daten hochladen (Excel oder DATEV-CSV)", type=["xlsx", "csv"])

if uploaded_file:
    with st.sidebar:
        st.markdown("### 📥 1. Daten-Quelle")
        mode = st.radio("Format auswählen", ["Standard Excel/CSV", "DATEV-Export (CSV)"])
        
        try:
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

            df_raw = df_raw.loc[:, ~df_raw.columns.str.contains('^Unnamed')].dropna(how='all', axis=0)
            all_cols = df_raw.columns.tolist()

            st.markdown("### 📍 2. Zuordnung")
            def find_idx(keys, default=0):
                for i, c in enumerate(all_cols):
                    if any(k.lower() in str(c).lower() for k in keys): return i
                return default if default < len(all_cols) else 0

            c_datum = st.selectbox("Rechnungsdatum", all_cols, index=find_idx(["datum", "belegdat"]))
            c_faellig = st.selectbox("Fälligkeitsdatum", all_cols, index=find_idx(["fällig", "termin"]))
            c_nr = st.selectbox("RE-Nummer", all_cols, index=find_idx(["belegfeld", "nummer", "re-nr"]))
            c_kunde = st.selectbox("Kunde", all_cols, index=find_idx(["name", "kunde"]))
            c_betrag = st.selectbox("Betrag (Brutto)", all_cols, index=find_idx(["brutto", "umsatz", "betrag"]))
            c_bezahlt = st.selectbox("Zahldatum", all_cols, index=find_idx(["bezahlt", "eingang"]))

            # --- DATEN-AUFBEREITUNG (CRASH-PROTECTION) ---
            df_work = df_raw.copy()
            
            # Konvertierung Betrag
            if df_work[c_betrag].dtype == 'object':
                df_work[c_betrag] = pd.to_numeric(df_work[c_betrag].astype(str).str.replace('.', '').str.replace(',', '.'), errors='coerce')
            
            # Speicherung der Original-Werte für "Fällig" (um "sofort" zu behalten)
            df_work[f"{c_faellig}_original"] = df_work[c_faellig].astype(str)
            
            # Konvertierung Daten
            df_work[c_datum] = pd.to_datetime(df_work[c_datum], errors='coerce')
            df_work[c_faellig] = pd.to_datetime(df_work[c_faellig], errors='coerce')
            df_work[c_bezahlt] = pd.to_datetime(df_work[c_bezahlt], errors='coerce')
            
            df_work = df_work.dropna(subset=[c_datum, c_betrag])

            st.markdown("### 🔍 3. Filter")
            if not df_work.empty:
                kunden_options = sorted(df_work[c_kunde].dropna().unique().tolist())
                sel_kunden = st.multiselect("Kunden auswählen", options=kunden_options, default=kunden_options)
                min_d, max_d = df_work[c_datum].min().date(), df_work[c_datum].max().date()
                date_range = st.date_input("Zeitraum", [min_d, max_d])

                st.markdown("---")
                start_btn = st.button("🚀 ANALYSE STARTEN", use_container_width=True)
            else:
                st.warning("Keine Daten gefunden. Prüfe Header & Spalten.")
                start_btn = False

        except Exception as e:
            st.error(f"Fehler bei der Spaltenzuordnung: {e}")
            start_btn = False

    # --- 3. HAUPTBEREICH ---
    if start_btn and len(date_range) == 2:
        mask = (df_work[c_datum].dt.date >= date_range[0]) & \
               (df_work[c_datum].dt.date <= date_range[1]) & \
               (df_work[c_kunde].isin(sel_kunden))
        f_df = df_work[mask].copy()

        if f_df.empty:
            st.warning("Keine Daten für die gewählten Filter gefunden.")
        else:
            tabs = st.tabs(["📊 Dashboard", "🔴 Mahnwesen", "📅 Cashflow", "💎 ABC/Risk", "🔍 Forensik"])

            with tabs[0]:
                m1, m2, m3 = st.columns(3)
                offen_mask = f_df[c_bezahlt].isna()
                m1.metric("Gesamtumsatz", format_euro(f_df[c_betrag].sum()))
                m2.metric("Offene Forderungen", format_euro(f_df[offen_mask][c_betrag].sum()))
                m3.metric("Belege", len(f_df))
                
                f_df['Monat'] = f_df[c_datum].dt.strftime('%Y-%m')
                chart = f_df.groupby('Monat')[c_betrag].sum().reset_index()
                st.plotly_chart(px.bar(chart, x='Monat', y=c_betrag, color_discrete_sequence=['#1E3A8A']), width='stretch')

            with tabs[1]:
                st.subheader("Fälligkeiten & Verzug")
                offen = f_df[offen_mask].copy()
                today = pd.Timestamp(datetime.now().date())
                
                # Verzug nur berechnen, wenn Fälligkeitsdatum existiert
                offen['Verzug'] = (today - offen[c_faellig]).dt.days
                
                # CRASH FIX: Konvertiere Spalten für Anzeige in Strings, um Arrow-Fehler zu vermeiden
                display_df = offen.copy()
                display_df[c_faellig] = display_df[f"{c_faellig}_original"] # Nutze Originalwert ("sofort")
                
                st.dataframe(display_df[[c_datum, c_faellig, c_kunde, c_betrag, 'Verzug']].sort_values(by='Verzug', ascending=False),
                             column_config={c_betrag: st.column_config.NumberColumn(format="%.2f €")}, 
                             width='stretch')
                
                st.download_button("Excel-Export", to_excel(offen), "Offene_Posten.xlsx")

            with tabs[2]:
                cf_data = f_df[f_df[c_bezahlt].isna()].groupby(c_faellig)[c_betrag].sum().reset_index()
                if not cf_data.empty:
                    st.plotly_chart(px.line(cf_data, x=c_faellig, y=c_betrag, title="Liquiditätsplanung", markers=True), width='stretch')

            with tabs[3]:
                abc = f_df.groupby(c_kunde)[c_betrag].sum().reset_index().sort_values(by=c_betrag, ascending=False)
                st.plotly_chart(px.pie(abc, values=c_betrag, names=c_kunde, hole=0.4), width='stretch')

            with tabs[4]:
                st.subheader("🔍 Forensik & Nummernkreis")
                try:
                    # Sicherer Check der Rechnungsnummern
                    raw_nums = pd.to_numeric(f_df[c_nr], errors='coerce').dropna().astype(int).sort_values().unique()
                    if len(raw_nums) > 1:
                        missing = np.setdiff1d(np.arange(raw_nums.min(), raw_nums.max() + 1), raw_nums)
                        if len(missing) > 0: st.warning(f"⚠️ {len(missing)} Nummern im Kreis fehlen.")
                        else: st.success("✅ Nummernkreis lückenlos.")
                except: st.info("Check nicht möglich.")
                
                err = f_df[f_df[c_bezahlt] < f_df[c_datum]]
                if not err.empty: st.error("❌ Zahlung vor Rechnungsdatum!"); st.dataframe(err, width='stretch')

    else:
        st.info("Bitte Datei laden und Filter einstellen. Dann 'Analyse starten'.")
