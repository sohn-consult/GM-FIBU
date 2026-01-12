import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO, StringIO
from datetime import datetime
import numpy as np

# --- 1. KONFIGURATION & DESIGN ---
st.set_page_config(
    page_title="Sohn-Consult | BI Dashboard",
    page_icon="👔",
    layout="wide"
)

# Clean Business Design CSS
st.markdown("""
    <style>
    .stApp { background-color: #FFFFFF; }
    [data-testid="stSidebar"] { background-color: #F8FAFC; border-right: 1px solid #E2E8F0; }
    [data-testid="stSidebar"] .stMarkdown p, 
    [data-testid="stSidebar"] label { color: #1E3A8A !important; font-weight: 600 !important; }
    h1, h2, h3 { color: #1E3A8A; font-family: 'Inter', sans-serif; }
    .stMetric { background-color: #FFFFFF; padding: 20px; border-radius: 10px; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1); border-top: 5px solid #1E3A8A; }
    .stTabs [aria-selected="true"] { background-color: #1E3A8A !important; color: white !important; border-radius: 4px; }
    .stButton>button { background-color: #1E3A8A; color: white; border-radius: 8px; font-weight: bold; height: 3em; width: 100%; }
    </style>
    """, unsafe_allow_html=True)

# Hilfsfunktion für deutsches Währungsformat (1.234,56 €)
def format_euro(val):
    if pd.isna(val): return "-"
    return f"{val:,.2f} €".replace(",", "X").replace(".", ",").replace("X", ".")

st.title("👔 Sohn-Consult | Strategische Fibu-Analyse")
st.caption("Professionelles Werkzeug für Performance-Reporting, Forensik & Liquiditätsplanung")
st.markdown("---")

# Hilfsfunktion für den Excel-Export
def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Sohn_Consult_Export')
    return output.getvalue()

# --- 2. DATEI-UPLOAD & IMPORT ---
uploaded_file = st.file_uploader("📂 Fibu-Datei hochladen (Excel oder DATEV-CSV)", type=["xlsx", "csv"])

if uploaded_file:
    with st.sidebar:
        st.markdown("### 📥 1. Datenquelle")
        mode = st.radio("Dateiformat", ["Excel / Standard CSV", "DATEV-Export (CSV)"])
        
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

            st.markdown("### 📍 2. Spalten-Mapping")
            def find_col(keys, default_idx):
                for i, col in enumerate(all_cols):
                    if any(k.lower() in str(col).lower() for k in keys): return i
                return default_idx if default_idx < len(all_cols) else 0

            c_datum = st.selectbox("Belegdatum", all_cols, index=find_col(["datum", "belegdat"], 2))
            c_faellig = st.selectbox("Fälligkeitsdatum", all_cols, index=find_col(["fällig", "termin"], 3))
            c_nr = st.selectbox("RE-Nummer", all_cols, index=find_col(["belegfeld", "nummer", "re-nr"], 1))
            c_kunde = st.selectbox("Kunde / Gegenkonto", all_cols, index=find_col(["name", "kunde", "gegenkonto"], 0))
            c_betrag = st.selectbox("Brutto-Betrag", all_cols, index=find_col(["brutto", "umsatz", "betrag"], 16))
            c_bezahlt = st.selectbox("Zahldatum (leer = offen)", all_cols, index=find_col(["bezahlt", "eingang", "ausgleich"], 17))

            # Transformation
            for col in [c_datum, c_faellig, c_bezahlt]:
                df[col] = pd.to_datetime(df[col], errors='coerce')
            
            if df[c_betrag].dtype == 'object':
                df[c_betrag] = pd.to_numeric(df[c_betrag].astype(str).str.replace('.', '').str.replace(',', '.'), errors='coerce')
            
            df = df.dropna(subset=[c_datum, c_betrag])

            st.markdown("### 🔍 3. Filter")
            alle_kunden = sorted(df[c_kunde].dropna().unique().tolist())
            sel_kunden = st.multiselect("Kunden auswählen", options=alle_kunden, default=alle_kunden)
            date_range = st.date_input("Zeitraum", [df[c_datum].min().date(), df[c_datum].max().date()])

            st.markdown("---")
            start_analysis = st.button("🚀 ANALYSE STARTEN", use_container_width=True)

        except Exception as e:
            st.error(f"Fehler beim Laden: {e}")
            start_analysis = False

    # --- 3. ANALYSE-OUTPUT ---
    if start_analysis and len(date_range) == 2:
        mask = (df[c_datum].dt.date >= date_range[0]) & (df[c_datum].dt.date <= date_range[1]) & (df[c_kunde].isin(sel_kunden))
        f_df = df[mask].copy()

        t1, t2, t3, t4, t5 = st.tabs(["📊 Performance", "🔴 Mahnwesen", "📅 Cashflow", "💎 ABC & Klumpen", "🔍 Forensik"])

        # KPI BERECHNUNGEN
        offen_mask = f_df[c_bezahlt].isna()
        total_rev = f_df[c_betrag].sum()
        open_sum = f_df[offen_mask][c_betrag].sum()

        with t1:
            col1, col2, col3 = st.columns(3)
            col1.metric("Gesamtumsatz", format_euro(total_rev))
            col2.metric("Offene Posten", format_euro(open_sum))
            col3.metric("Anzahl Belege", len(f_df))
            
            f_df['Monat'] = f_df[c_datum].dt.strftime('%Y-%m')
            monats_chart = f_df.groupby('Monat')[c_betrag].sum().reset_index()
            fig = px.bar(monats_chart, x='Monat', y=c_betrag, title="Umsatzentwicklung", 
                         color_discrete_sequence=['#1E3A8A'], labels={c_betrag: "Umsatz in €"})
            fig.update_layout(yaxis_tickformat=',.2f') # Plotly Format
            st.plotly_chart(fig, use_container_width=True)

        with t2:
            st.subheader("Mahnwesen (Aging)")
            offene_df = f_df[offen_mask].copy()
            offene_df['Verzug (Tage)'] = (pd.Timestamp(datetime.now().date()) - offene_df[c_faellig]).dt.days
            
            # Formatiere Tabelle für die Anzeige
            display_op = offene_df[[c_datum, c_faellig, c_kunde, c_betrag, 'Verzug (Tage)']].copy()
            st.dataframe(display_op.sort_values(by='Verzug (Tage)', ascending=False), 
                         column_config={c_betrag: st.column_config.NumberColumn(format="%.2f €")},
                         use_container_width=True)
            st.download_button("Excel-Export (Offene Posten)", to_excel(offene_df), "Offene_Posten.xlsx")

        with t3:
            st.subheader("Cashflow-Prognose")
            cf = offene_df.groupby(c_faellig)[c_betrag].sum().reset_index()
            if not cf.empty:
                fig_cf = px.line(cf, x=c_faellig, y=c_betrag, markers=True, title="Liquiditätsplanung", color_discrete_sequence=['#10B981'])
                fig_cf.update_layout(yaxis_tickformat=',.2f')
                st.plotly_chart(fig_cf, use_container_width=True)

        with t4:
            st.subheader("ABC-Analyse")
            abc = f_df.groupby(c_kunde)[c_betrag].sum().reset_index().sort_values(by=c_betrag, ascending=False)
            st.dataframe(abc, column_config={c_betrag: st.column_config.NumberColumn(format="%.2f €")}, use_container_width=True)

        with t5:
            st.subheader("🔍 Forensik-Modul")
            # Nummernkreis-Lücken
            try:
                nums = pd.to_numeric(f_df[c_nr], errors='coerce').dropna().astype(int).sort_values().unique()
                if len(nums) > 1:
                    missing = np.setdiff1d(np.arange(nums.min(), nums.max() + 1), nums)
                    if len(missing) > 0: st.warning(f"⚠️ Nummernkreis-Lücke: {len(missing)} Nummern fehlen.")
                    else: st.success("✅ RE-Nummern lückenlos.")
            except: st.info("Nummernkreis-Check nicht möglich.")

            # Betrags-Forensik (Ausreißer)
            q_high = f_df[c_betrag].quantile(0.95)
            st.info(f"Top 5% der Rechnungen (> {format_euro(q_high)}):")
            st.dataframe(f_df[f_df[c_betrag] > q_high][[c_datum, c_kunde, c_betrag]], 
                         column_config={c_betrag: st.column_config.NumberColumn(format="%.2f €")}, use_container_width=True)

    else:
        st.info("👋 Willkommen! Bitte laden Sie eine Datei hoch und klicken Sie auf 'ANALYSE STARTEN'.")
