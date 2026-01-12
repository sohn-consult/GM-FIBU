import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO, StringIO
from datetime import datetime

# --- 1. DESIGN & KONFIGURATION ---
st.set_page_config(page_title="Sohn-Consult Business Intelligence", page_icon="📈", layout="wide")

# Custom CSS für das Sohn-Consult Corporate Design
st.markdown("""
    <style>
    .stApp { background-color: #F8FAFC; }
    [data-testid="stSidebar"] { background-color: #1E3A8A; color: white; }
    [data-testid="stSidebar"] .stMarkdown p { color: white; }
    h1, h2, h3 { color: #1E3A8A; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
    .stMetric { background-color: white; padding: 20px; border-radius: 12px; box-shadow: 0 4px 12px rgba(0,0,0,0.05); border: 1px solid #E2E8F0; }
    div[data-testid="stMetricValue"] { color: #1E3A8A; }
    .stButton>button { background-color: #1E3A8A; color: white; width: 100%; border-radius: 8px; font-weight: bold; border: none; padding: 10px; }
    .stButton>button:hover { background-color: #2563EB; border: none; color: white; }
    .stTabs [data-baseweb="tab-list"] { gap: 24px; }
    .stTabs [data-baseweb="tab"] { height: 50px; white-space: pre-wrap; background-color: transparent; border-radius: 4px 4px 0 0; gap: 1px; padding-top: 10px; }
    </style>
    """, unsafe_allow_html=True)

# Header
st.title("📊 Sohn-Consult | Universal Fibu-Analyse")
st.caption("Strategisches Cockpit für Debitorenmanagement, Cashflow-Prognose & Forensik")
st.markdown("---")

# Hilfsfunktion für Excel-Export
def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Sohn_Consult_Analyse')
    return output.getvalue()

# --- 2. DATEI-UPLOAD ---
uploaded_file = st.file_uploader("📂 Fibu-Datei hochladen (Excel oder DATEV-CSV)", type=["xlsx", "csv"])

if uploaded_file:
    with st.sidebar:
        st.markdown("### ⚙️ 1. Import-Modus")
        mode = st.radio("Dateiformat wählen", ["Standard Excel/CSV", "DATEV-Export (CSV)"])
        
        try:
            # Einlesen der Datei
            if mode == "DATEV-Export (CSV)":
                raw_bytes = uploaded_file.getvalue()
                content = raw_bytes.decode('latin-1') # DATEV Standard
                # Header-Suche: DATEV-Dateien haben oft Metadaten in Zeile 1
                df = pd.read_csv(StringIO(content), sep=None, engine='python', skiprows=1)
                if "Umsatz" not in df.columns: # Zweiter Versuch falls Header anders
                    df = pd.read_csv(StringIO(content), sep=None, engine='python')
            else:
                if uploaded_file.name.endswith('.csv'):
                    df = pd.read_csv(uploaded_file, sep=None, engine='python')
                else:
                    header_row = st.number_input("Header-Zeile (Spaltennamen)", min_value=1, value=3)
                    xl = pd.ExcelFile(uploaded_file)
                    sheet = st.selectbox("Tabellenblatt", xl.sheet_names)
                    df = pd.read_excel(uploaded_file, sheet_name=sheet, header=header_row-1)

            # Datenreinigung
            df = df.loc[:, ~df.columns.str.contains('^Unnamed')].dropna(how='all', axis=0)
            all_cols = df.columns.tolist()

            st.markdown("### 📍 2. Spalten-Mapping")
            def find_col(keys, default_idx):
                for i, col in enumerate(all_cols):
                    if any(k.lower() in str(col).lower() for k in keys): return i
                return default_idx if default_idx < len(all_cols) else 0

            c_datum = st.selectbox("Rechnungsdatum", all_cols, index=find_col(["datum", "belegdat"], 2))
            c_faellig = st.selectbox("Fälligkeitsdatum", all_cols, index=find_col(["fällig", "skonto", "termin"], 3))
            c_nr = st.selectbox("RE-Nummer", all_cols, index=find_col(["belegfeld", "nummer", "re-nr"], 1))
            c_kunde = st.selectbox("Kunde / Projekt", all_cols, index=find_col(["name", "kunde", "gegenkonto"], 0))
            c_betrag = st.selectbox("Betrag (Brutto)", all_cols, index=find_col(["brutto", "umsatz", "betrag"], 16))
            c_bezahlt = st.selectbox("Zahlungseingang (Datum)", all_cols, index=find_col(["bezahlt", "eingang", "ausgleich"], 17))

            # Transformation
            df[c_datum] = pd.to_datetime(df[c_datum], errors='coerce')
            df[c_faellig] = pd.to_datetime(df[c_faellig], errors='coerce')
            df[c_bezahlt] = pd.to_datetime(df[c_bezahlt], errors='coerce')
            if df[c_betrag].dtype == 'object':
                df[c_betrag] = pd.to_numeric(df[c_betrag].astype(str).str.replace('.', '').str.replace(',', '.'), errors='coerce')
            
            df = df.dropna(subset=[c_datum, c_betrag])

            st.markdown("### 🔍 3. Filter")
            kunden = sorted(df[c_kunde].dropna().unique().tolist())
            sel_kunden = st.multiselect("Kunden filtern", options=kunden, default=kunden)
            
            min_date, max_date = df[c_datum].min().date(), df[c_datum].max().date()
            date_range = st.date_input("Zeitraum", [min_date, max_date])

            st.markdown("---")
            start_btn = st.button("🚀 ANALYSE STARTEN")

        except Exception as e:
            st.error(f"Fehler beim Laden der Datei: {e}")
            start_btn = False

    # --- 3. HAUPTBEREICH: ANALYSE ---
    if start_btn and len(date_range) == 2:
        start_d, end_d = date_range
        mask = (df[c_datum].dt.date >= start_d) & (df[c_datum].dt.date <= end_d) & (df[c_kunde].isin(sel_kunden))
        working_df = df[mask].copy()

        # Tabs
        tab_main, tab_aging, tab_cash, tab_abc, tab_forensic = st.tabs([
            "📊 Dashboard", "🔴 Mahnwesen", "📅 Cashflow", "💎 ABC-Analyse", "🔍 Forensik"
        ])

        # --- KPI BERECHNUNG ---
        offene_mask = working_df[c_bezahlt].isna()
        total_rev = working_df[c_betrag].sum()
        open_sum = working_df[offene_mask][c_betrag].sum()
        
        with tab_main:
            col1, col2, col3 = st.columns(3)
            with col1: st.metric("Gesamtumsatz", f"{total_rev:,.2f} €")
            with col2: st.metric("Offene Forderungen", f"{open_sum:,.2f} €", delta=f"{(open_sum/total_rev*100):.1f}% Quote", delta_color="inverse")
            with col3:
                # Zahlungshistorie
                paid_df = working_df[working_df[c_bezahlt].notna()]
                avg_days = (paid_df[c_bezahlt] - paid_df[c_datum]).dt.days.mean()
                st.metric("Ø Zahlungsdauer", f"{avg_days:.1f} Tage" if not pd.isna(avg_days) else "N/A")

            working_df['Monat'] = working_df[c_datum].dt.strftime('%Y-%m')
            monats_chart = working_df.groupby('Monat')[c_betrag].sum().reset_index()
            st.plotly_chart(px.bar(monats_chart, x='Monat', y=c_betrag, title="Umsatz nach Monat", color_discrete_sequence=['#1E3A8A']), use_container_width=True)

        with tab_aging:
            st.subheader("Mahnwesen & Aging")
            offene_df = working_df[offene_mask].copy()
            today = pd.Timestamp(datetime.now().date())
            offene_df['Verzug (Tage)'] = (today - offene_df[c_faellig]).dt.days
            
            def color_aging(val):
                if val > 30: return 'background-color: #FEE2E2; color: #991B1B; font-weight: bold;'
                if val > 14: return 'background-color: #FFEDD5; color: #9A3412;'
                return ''
            
            if not offene_df.empty:
                st.dataframe(offene_df[[c_datum, c_faellig, c_kunde, c_betrag, 'Verzug (Tage)']].sort_values(by='Verzug (Tage)', ascending=False).style.applymap(color_aging, subset=['Verzug (Tage)']), use_container_width=True)
                st.download_button("Excel-Liste Offene Posten", to_excel(offene_df), "Sohn_Consult_OP_Liste.xlsx")
            else:
                st.success("Alle Rechnungen im gewählten Filter sind bezahlt!")

        with tab_cash:
            st.subheader("Cashflow-Prognose (Erwartete Eingänge)")
            cf_df = offene_df.groupby(c_faellig)[c_betrag].sum().reset_index()
            if not cf_df.empty:
                st.plotly_chart(px.line(cf_df, x=c_faellig, y=c_betrag, markers=True, title="Liquiditätszufluss nach Fälligkeitsdatum", color_discrete_sequence=['#10B981']), use_container_width=True)
            else:
                st.info("Keine Daten für Prognose vorhanden.")

        with tab_abc:
            st.subheader("Strategische ABC-Analyse")
            abc = working_df.groupby(c_kunde)[c_betrag].sum().reset_index().sort_values(by=c_betrag, ascending=False)
            abc['CumSum'] = abc[c_betrag].cumsum()
            abc['Pct'] = (abc['CumSum'] / abc[c_betrag].sum()) * 100
            abc['Kategorie'] = abc['Pct'].apply(lambda x: 'A (Top 80%)' if x <= 80 else ('B (80-95%)' if x <= 95 else 'C (Rest)'))
            
            st.plotly_chart(px.pie(abc, values=c_betrag, names='Kategorie', hole=0.4, title="Umsatzanteile nach Kategorien", color_discrete_sequence=['#1E3A8A', '#3B82F6', '#93C5FD']), use_container_width=True)
            st.dataframe(abc[[c_kunde, c_betrag, 'Kategorie']], use_container_width=True)

        with tab_forensic:
            st.subheader("🔍 Forensischer Check & Integrität")
            col_a, col_b = st.columns(2)
            
            with col_a:
                # Dubletten
                dubs = working_df[working_df.duplicated(subset=[c_nr], keep=False)]
                if not dubs.empty:
                    st.warning(f"Achtung: {len(dubs)} Dubletten bei Rechnungsnummern gefunden!")
                    st.dataframe(dubs)
                else:
                    st.success("Keine RE-Nummern Dubletten.")

            with col_b:
                # Logikfehler
                logik = working_df[working_df[c_bezahlt] < working_df[c_datum]]
                if not logik.empty:
                    st.error(f"Kritisch: {len(logik)} Zahlungen vor Rechnungsdatum!")
                    st.dataframe(logik)
                else:
                    st.success("Datum-Logik ist konsistent.")

            # Ausreißer
            q_high = working_df[c_betrag].quantile(0.95)
            outliers = working_df[working_df[c_betrag] > q_high]
            st.info(f"Auffällig hohe Rechnungen (Top 5% > {q_high:,.2f} €):")
            st.dataframe(outliers)

    else:
        # Willkommens-Bildschirm
        st.info("👋 Willkommen! Bitte laden Sie links eine Datei hoch und konfigurieren Sie die Spalten, um die Analyse zu starten.")
