import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO, StringIO
from datetime import datetime
import numpy as np

# --- 1. DESIGN & KONFIGURATION ---
st.set_page_config(
    page_title="Sohn-Consult | BI Dashboard",
    page_icon="👔",
    layout="wide"
)

# Custom CSS für professionelles Branding und maximale Lesbarkeit
st.markdown("""
    <style>
    .stApp { background-color: #FFFFFF; }
    [data-testid="stSidebar"] {
        background-color: #F1F5F9; 
        border-right: 1px solid #CBD5E1;
    }
    [data-testid="stSidebar"] .stMarkdown p, 
    [data-testid="stSidebar"] label {
        color: #0F172A !important;
        font-weight: 600 !important;
    }
    h1, h2, h3 { color: #1E3A8A; font-family: 'Inter', sans-serif; font-weight: 700; }
    .stMetric {
        background-color: #FFFFFF;
        padding: 20px;
        border-radius: 10px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.05);
        border-top: 5px solid #1E3A8A;
    }
    .stButton>button {
        background-color: #1E3A8A;
        color: white;
        border-radius: 8px;
        font-weight: bold;
        width: 100%;
        height: 3.5em;
    }
    .stTabs [aria-selected="true"] {
        background-color: #1E3A8A !important;
        color: white !important;
        border-radius: 4px;
    }
    </style>
    """, unsafe_allow_html=True)

# Hilfsfunktion für Euro-Formatierung (1.234,56 €)
def format_euro(val):
    if pd.isna(val): return "0,00 €"
    return f"{val:,.2f} €".replace(",", "X").replace(".", ",").replace("X", ".")

# Hilfsfunktion für den Excel-Export
def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Sohn_Consult_Analyse')
    return output.getvalue()

st.title("👔 Sohn-Consult | Business Intelligence")
st.caption("Professionelles Reporting für Fibu-Daten, Forensik & Cashflow")
st.markdown("---")

# --- 2. DATEI-UPLOAD ---
uploaded_file = st.file_uploader("📂 Fibu-Daten hochladen (Excel oder DATEV-CSV)", type=["xlsx", "csv"])

if uploaded_file:
    with st.sidebar:
        st.markdown("### 📥 1. Daten-Quelle")
        mode = st.radio("Format auswählen", ["Standard Excel/CSV", "DATEV-Export (CSV)"])
        
        try:
            # Einlese-Logik je nach Format
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

            # Basis-Reinigung: Entferne leere Spalten/Zeilen
            df_raw = df_raw.loc[:, ~df_raw.columns.str.contains('^Unnamed')].dropna(how='all', axis=0)
            all_cols = df_raw.columns.tolist()

            st.markdown("### 📍 2. Spalten-Zuordnung")
            def find_idx(keys, default=0):
                for i, c in enumerate(all_cols):
                    if any(k.lower() in str(c).lower() for k in keys): return i
                return default if default < len(all_cols) else 0

            # Mapping der Spalten
            c_datum = st.selectbox("Rechnungsdatum", all_cols, index=find_idx(["datum", "belegdat"]))
            c_faellig = st.selectbox("Fälligkeitsdatum", all_cols, index=find_idx(["fällig", "termin"]))
            c_nr = st.selectbox("RE-Nummer", all_cols, index=find_idx(["belegfeld", "nummer", "re-nr"]))
            c_kunde = st.selectbox("Kunde", all_cols, index=find_idx(["name", "kunde"]))
            c_betrag = st.selectbox("Betrag (Brutto)", all_cols, index=find_idx(["brutto", "umsatz", "betrag"]))
            c_bezahlt = st.selectbox("Zahldatum (leer = offen)", all_cols, index=find_idx(["bezahlt", "eingang"]))

            # --- DATEN-TRANSFORMATION ---
            df_clean = df_raw.copy()
            for col in [c_datum, c_faellig, c_bezahlt]:
                df_clean[col] = pd.to_datetime(df_clean[col], errors='coerce')
            
            if df_clean[c_betrag].dtype == 'object':
                df_clean[c_betrag] = pd.to_numeric(df_clean[c_betrag].astype(str).str.replace('.', '').str.replace(',', '.'), errors='coerce')
            
            # Wichtig: Nur Zeilen mit Datum und Betrag behalten
            df_clean = df_clean.dropna(subset=[c_datum, c_betrag])

            # --- FILTER ---
            st.markdown("### 🔍 3. Filter")
            if not df_clean.empty:
                kunden_liste = sorted(df_clean[c_kunde].dropna().unique().tolist())
                sel_kunden = st.multiselect("Kunden filtern", options=kunden_liste, default=kunden_liste)
                
                min_d, max_d = df_clean[c_datum].min().date(), df_clean[c_datum].max().date()
                date_range = st.date_input("Zeitraum wählen", [min_d, max_d])

                st.markdown("---")
                start_btn = st.button("🚀 ANALYSE STARTEN", use_container_width=True)
            else:
                st.warning("Keine gültigen Daten gefunden. Prüfe die Header-Zeile.")
                start_btn = False

        except Exception as e:
            st.error(f"Fehler beim Einlesen: {e}")
            start_btn = False

    # --- 3. HAUPTBEREICH: ANALYSE ---
    if start_btn and len(date_range) == 2:
        # Filteranwendung
        mask = (df_clean[c_datum].dt.date >= date_range[0]) & \
               (df_clean[c_datum].dt.date <= date_range[1]) & \
               (df_clean[c_kunde].isin(sel_kunden))
        f_df = df_clean[mask].copy()

        if f_df.empty:
            st.warning("Keine Daten für die gewählten Filter vorhanden.")
        else:
            tabs = st.tabs(["📊 Dashboard", "🔴 Mahnwesen", "📅 Cashflow", "💎 ABC/Risk", "🔍 Forensik"])

            # --- TAB 1: DASHBOARD ---
            with tabs[0]:
                m1, m2, m3 = st.columns(3)
                offen_mask = f_df[c_bezahlt].isna()
                total_rev = f_df[c_betrag].sum()
                open_sum = f_df[offen_mask][c_betrag].sum()
                
                m1.metric("Gesamtumsatz", format_euro(total_rev))
                m2.metric("Offene Forderungen", format_euro(open_sum))
                m3.metric("Anzahl Belege", len(f_df))
                
                f_df['Monat'] = f_df[c_datum].dt.strftime('%Y-%m')
                monats_data = f_df.groupby('Monat')[c_betrag].sum().reset_index()
                fig = px.bar(monats_data, x='Monat', y=c_betrag, labels={c_betrag: "Umsatz in €"}, 
                             color_discrete_sequence=['#1E3A8A'], title="Umsatzentwicklung")
                fig.update_layout(yaxis_tickformat=',.2f')
                st.plotly_chart(fig, use_container_width=True)

            # --- TAB 2: MAHNWESEN ---
            with tabs[1]:
                st.subheader("Übersicht überfälliger Rechnungen")
                offen = f_df[offen_mask].copy()
                today = pd.Timestamp(datetime.now().date())
                offen['Verzug'] = (today - offen[c_faellig]).dt.days
                
                if not offen.empty:
                    st.dataframe(offen[[c_datum, c_faellig, c_kunde, c_betrag, 'Verzug']].sort_values(by='Verzug', ascending=False), 
                                 column_config={c_betrag: st.column_config.NumberColumn(format="%.2f €")}, 
                                 use_container_width=True)
                    st.download_button("📥 Offene Posten (Excel)", to_excel(offen), "Offene_Posten_SohnConsult.xlsx")
                else:
                    st.success("Hervorragend! Alle Forderungen im Zeitraum sind beglichen.")

            # --- TAB 3: CASHFLOW ---
            with tabs[2]:
                st.subheader("Liquiditätsprognose")
                cf_data = offen.groupby(c_faellig)[c_betrag].sum().reset_index()
                if not cf_data.empty:
                    fig_cf = px.line(cf_data, x=c_faellig, y=c_betrag, markers=True, 
                                     title="Erwartete Zahlungseingänge nach Fälligkeit", color_discrete_sequence=['#10B981'])
                    fig_cf.update_layout(yaxis_tickformat=',.2f')
                    st.plotly_chart(fig_cf, use_container_width=True)
                else:
                    st.info("Keine offenen Rechnungen für eine Cashflow-Prognose verfügbar.")

            # --- TAB 4: ABC & KLUMPENRISIKO ---
            with tabs[3]:
                st.subheader("Strategische Kundenanalyse")
                abc = f_df.groupby(c_kunde)[c_betrag].sum().reset_index().sort_values(by=c_betrag, ascending=False)
                st.plotly_chart(px.pie(abc, values=c_betrag, names=c_kunde, hole=0.4, title="Umsatzverteilung nach Kunden"), use_container_width=True)
                
                top_3_pct = (abc[c_betrag].head(3).sum() / abc[c_betrag].sum()) * 100
                st.metric("Klumpenrisiko (Top 3 Kunden)", f"{top_3_pct:.1f}%")
                if top_3_pct > 60: st.warning("Strategischer Hinweis: Sehr hohe Abhängigkeit von wenigen Auftraggebern.")

            # --- TAB 5: FORENSIK ---
            with tabs[4]:
                st.subheader("Forensische Prüfung & Daten-Integrität")
                
                # 1. Nummernkreis-Check
                try:
                    nums = pd.to_numeric(f_df[c_nr], errors='coerce').dropna().astype(int).sort_values().unique()
                    if len(nums) > 1:
                        missing = np.setdiff1d(np.arange(nums.min(), nums.max() + 1), nums)
                        if len(missing) > 0:
                            st.warning(f"⚠️ Nummernkreis-Lücke: {len(missing)} Nummern fehlen (z.B. {missing[:5]}...)")
                        else: st.success("✅ Rechnungsnummern sind lückenlos.")
                except: st.info("Prüfung der RE-Nummern für dieses Format nicht möglich.")
                
                # 2. Logik-Check
                logik_err = f_df[f_df[c_bezahlt] < f_df[c_datum]]
                if not logik_err.empty:
                    st.error("❌ Kritisch: Zahlung vor Rechnungsdatum gefunden!")
                    st.dataframe(logik_err[[c_datum, c_bezahlt, c_kunde, c_betrag]])
                
                # 3. Betrags-Check
                q_high = f_df[c_betrag].quantile(0.95)
                st.info(f"💡 Statistische Ausreißer (Top 5% > {format_euro(q_high)}):")
                st.dataframe(f_df[f_df[c_betrag] > q_high][[c_datum, c_kunde, c_betrag]], 
                             column_config={c_betrag: st.column_config.NumberColumn(format="%.2f €")})

    else:
        st.info("👋 Willkommen! Bitte laden Sie eine Datei hoch und konfigurieren Sie die Spalten in der Sidebar.")
        if 'df_raw' in locals():
            st.markdown("### 📄 Vorschau der Rohdaten (erste 5 Zeilen):")
            st.dataframe(df_raw.head(5))
