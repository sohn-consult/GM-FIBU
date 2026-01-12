import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO, StringIO
from datetime import datetime
import numpy as np
import re

# --- 1. CONFIG & MODERN UI ---
st.set_page_config(page_title="Sohn-Consult | Executive BI", page_icon="👔", layout="wide")

# Modernes High-Contrast Design für Berater
st.markdown("""
    <style>
    .stApp { background-color: #F8FAFC; }
    [data-testid="stSidebar"] { background-color: #F1F5F9; border-right: 1px solid #CBD5E1; }
    h1, h2, h3 { color: #1E3A8A; font-family: 'Inter', sans-serif; font-weight: 700; }
    .stMetric { 
        background-color: #FFFFFF; padding: 20px; border-radius: 12px; 
        box-shadow: 0 4px 6px rgba(0,0,0,0.05); border-top: 5px solid #1E3A8A; 
    }
    .stButton>button { 
        background-color: #1E3A8A; color: white; border-radius: 8px; 
        font-weight: bold; width: 100%; height: 3.5em; 
    }
    .stTabs [aria-selected="true"] { background-color: #1E3A8A !important; color: white !important; }
    </style>
    """, unsafe_allow_html=True)

# --- 2. CORE FUNCTIONS (STABILITY LAYER) ---
def format_euro(val):
    if pd.isna(val) or val is None: return "0,00 €"
    return f"{val:,.2f} €".replace(",", "X").replace(".", ",").replace("X", ".")

def clean_dataframe(df):
    """Verhindert TypeError und Arrow-Abstürze durch Bereinigung beim Import."""
    # Alle Spaltennamen zu bereinigten Strings machen
    df.columns = [str(c).strip() for c in df.columns]
    # Entferne 'Unnamed' Spalten sicher ohne mathematische Operatoren
    valid_cols = [c for c in df.columns if "Unnamed" not in c and c.lower() != "nan" and c != ""]
    df = df[valid_cols].copy()
    return df.dropna(how='all', axis=0)

def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='SohnConsult_Export')
    return output.getvalue()

# --- 3. APP HEADER ---
st.title("👔 Sohn-Consult | Strategic BI Dashboard")
st.caption("Professionelle Analyse: Performance, Forensic, Bank & Cashflow - Version 2026.6 (Bulletproof)")
st.markdown("---")

# --- 4. DATA IMPORT ---
col_u1, col_u2 = st.columns(2)
with col_u1:
    fibu_file = st.file_uploader("📂 1. Fibu-Datei laden (XLSX/CSV)", type=["xlsx", "csv"])
with col_u2:
    bank_file = st.file_uploader("🏦 2. OPTIONAL: Bank-CSV laden", type=["csv"])

if fibu_file:
    with st.sidebar:
        st.header("⚙️ Konfiguration")
        mode = st.radio("Format", ["Standard Excel/CSV", "DATEV Export"])
        header_row = st.number_input("Header-Zeile", min_value=1, value=3)
        
        try:
            if mode == "DATEV Export":
                content = fibu_file.getvalue().decode('latin-1', errors='ignore')
                df_raw = pd.read_csv(StringIO(content), sep=None, engine='python', skiprows=1)
            else:
                if fibu_file.name.endswith('.csv'):
                    df_raw = pd.read_csv(fibu_file, sep=None, engine='python')
                else:
                    df_raw = pd.read_excel(fibu_file, header=int(header_row-1))
            
            # Sanity Check & Reinigung
            df_work = clean_dataframe(df_raw)
            cols = df_work.columns.tolist()

            st.subheader("📍 Mapping")
            def find_idx(keys):
                for i, c in enumerate(cols):
                    if any(k.lower() in str(c).lower() for k in keys): return i
                return 0

            c_dat = st.selectbox("Rechnungsdatum", cols, index=find_idx(["datum", "belegdat"]))
            c_fae = st.selectbox("Fälligkeitsdatum", cols, index=find_idx(["fällig", "termin"]))
            c_nr = st.selectbox("RE-Nummer", cols, index=find_idx(["nummer", "belegfeld"]))
            c_kun = st.selectbox("Kunde", cols, index=find_idx(["kunde", "name"]))
            c_bet = st.selectbox("Betrag", cols, index=find_idx(["brutto", "betrag", "umsatz"]))
            c_pay = st.selectbox("Zahldatum", cols, index=find_idx(["gezahlt", "ausgleich", "eingang"]))

            # Transformationen mit Fehlerunterdrückung
            df_work[c_dat] = pd.to_datetime(df_work[c_dat], errors='coerce')
            # Für die Anzeige speichern wir Originale als Strings (Arrow-Fix)
            df_work['Fällig_Anzeige'] = df_work[c_fae].astype(str).fillna("unbekannt")
            df_work[c_fae] = pd.to_datetime(df_work[c_fae], errors='coerce')
            df_work[c_pay] = pd.to_datetime(df_work[c_pay], errors='coerce')
            
            if df_work[c_bet].dtype == 'object':
                df_work[c_bet] = pd.to_numeric(df_work[c_bet].astype(str).str.replace('.', '').str.replace(',', '.'), errors='coerce')
            
            df_work = df_work.dropna(subset=[c_dat, c_bet])

            st.markdown("### 🔍 Filter")
            k_list = sorted(df_work[c_kun].dropna().unique().tolist())
            sel_kunden = st.multiselect("Kunden-Fokus", options=k_list, default=k_list)
            
            # Zeitraum sicherstellen
            min_date = df_work[c_dat].min().date()
            max_date = df_work[c_dat].max().date()
            date_range = st.date_input("Analyse-Zeitraum", [min_date, max_date])

            # Finaler Start-Befehl
            start_btn = st.button("🚀 ANALYSE STARTEN", width='stretch')

        except Exception as e:
            st.error(f"Fehler bei Datenaufbereitung: {e}")
            start_btn = False

    # --- 5. EXECUTION ENGINE ---
    if start_btn and len(date_range) == 2:
        mask = (df_work[c_dat].dt.date >= date_range[0]) & \
               (df_work[c_dat].dt.date <= date_range[1]) & \
               (df_work[c_kun].isin(sel_kunden))
        f_df = df_work[mask].copy()

        today = pd.Timestamp(datetime.now().date())
        df_offen = f_df[f_df[c_pay].isna()].copy()
        df_paid = f_df[~f_df[c_pay].isna()].copy()

        tabs = st.tabs(["📊 Performance", "🔴 Aging & OP", "💎 Strategie", "🔍 Forensik", "🏦 Bank-Match"])

        with tabs[0]:
            # KPIs
            k1, k2, k3, k4 = st.columns(4)
            rev_total = f_df[c_bet].sum()
            k1.metric("Gesamtumsatz", format_euro(rev_total))
            k2.metric("Offene Posten", format_euro(df_offen[c_bet].sum()))
            
            dso = (df_paid[c_pay] - df_paid[c_dat]).dt.days.mean() if not df_paid.empty else 0
            k3.metric("Ø Zahlungszeit (DSO)", f"{dso:.1f} Tage" if dso > 0 else "N/A")
            k4.metric("Belege", len(f_df))

            # Charts
            cp1, cp2 = st.columns([2, 1])
            with cp1:
                f_df['Monat'] = f_df[c_dat].dt.strftime('%Y-%m')
                monat_sum = f_df.groupby('Monat')[c_bet].sum().reset_index()
                st.plotly_chart(px.bar(monat_sum, x='Monat', y=c_bet, color_discrete_sequence=['#1E3A8A'], title="Umsatz nach Monat"), width='stretch')
            with cp2:
                f_df = f_df.sort_values(c_dat)
                f_df['Kumuliert'] = f_df[c_bet].cumsum()
                st.plotly_chart(px.area(f_df, x=c_dat, y='Kumuliert', title="Jahres-Wachstumspfad", color_discrete_sequence=['#3B82F6']), width='stretch')

        with tabs[1]:
            st.subheader("Forderungs-Management & Liquiditätsvorschau")
            ca1, ca2 = st.columns([1, 2])
            with ca1:
                df_offen['Verzug'] = (today - df_offen[c_fae]).dt.days
                def bucket_func(d):
                    if pd.isna(d): return "5. Unbekannt"
                    if d <= 0: return "1. Pünktlich"
                    if d <= 30: return "2. 1-30 Tage"
                    if d <= 60: return "3. 31-60 Tage"
                    return "4. > 60 Tage"
                df_offen['Bucket'] = df_offen['Verzug'].apply(bucket_func)
                st.plotly_chart(px.pie(df_offen.groupby('Bucket')[c_bet].sum().reset_index(), values=c_bet, names='Bucket', hole=0.5, title="Risikoprofil"), width='stretch')
            
            with ca2:
                # STABILITÄTS-FIX: Absoluter Wert für Scatter-Punktgröße gegen ValueError
                df_predict = df_offen.groupby(c_fae)[c_bet].sum().reset_index()
                if not df_predict.empty:
                    df_predict['Betrag_Safe'] = df_predict[c_bet].abs().clip(lower=0.1)
                    st.plotly_chart(px.scatter(df_predict, x=c_fae, y=c_bet, size='Betrag_Safe', title="Cash-Inflow Prognose", color_discrete_sequence=['#10B981']), width='stretch')
            
            # ARROW-FIX: Daten für Anzeige als Text kopieren
            disp_op = df_offen[[c_dat, c_fae, c_kun, c_bet, 'Verzug']].copy()
            disp_op[c_dat] = disp_op[c_dat].dt.strftime('%d.%m.%Y')
            disp_op[c_fae] = df_offen['Fällig_Anzeige']
            st.dataframe(disp_op.sort_values('Verzug', ascending=False), column_config={c_bet: st.column_config.NumberColumn(format="%.2f €")}, width='stretch')
            st.download_button("📥 Offene Posten Excel", to_excel(df_offen), "SohnConsult_OP_Liste.xlsx")

        with tabs[2]:
            st.subheader("Strategische Analyse (ABC)")
            abc_data = f_df.groupby(c_kun)[c_bet].sum().reset_index().sort_values(by=c_bet, ascending=False)
            st.plotly_chart(px.bar(abc_data.head(15), x=c_kun, y=c_bet, color_discrete_sequence=['#1E3A8A'], title="Top 15 Kundenumsätze"), width='stretch')
            top3_val = (abc_data[c_bet].head(3).sum() / rev_total * 100) if rev_total > 0 else 0
            st.metric("Klumpenrisiko (Top 3)", f"{top3_val:.1f}%")

        with tabs[3]:
            st.subheader("🔍 Daten-Forensik & Integrität")
            miss = f_df.isna().sum().sum()
            if miss == 0: st.success("✅ Datensatz ist vollständig.")
            else: st.info(f"ℹ️ {miss} leere Felder im Datensatz.")

            fl1, fl2 = st.columns(2)
            with fl1:
                st.markdown("### Plausibilitäts-Prüfung")
                err_logik = f_df[f_df[c_pay] < f_df[c_dat]]
                if not err_logik.empty:
                    st.error(f"❌ {len(err_logik)} Zahlungen VOR Rechnungsdatum gefunden.")
                    st.dataframe(err_logik[[c_dat, c_pay, c_kun, c_bet]])
                else: st.success("✅ Zeitliche Abfolge der Zahlungen ist logisch.")
            with fl2:
                st.markdown("### Nummernkreis-Prüfung")
                try:
                    def ext_n(v):
                        found = re.findall(r'\d+', str(v))
                        return int(found[-1]) if found else None
                    nums = f_df[c_nr].apply(ext_n).dropna().unique()
                    if len(nums) > 1:
                        full_range = np.arange(nums.min(), nums.max() + 1)
                        missing = np.setdiff1d(full_range, nums)
                        if len(missing) > 0:
                            st.warning(f"⚠️ {len(missing)} Rechnungsnummern fehlen im Kreis.")
                            with st.expander("Details anzeigen"): st.write(missing[:50])
                        else: st.success("✅ Rechnungsnummernkreis ist lückenlos.")
                except: st.info("Check für Nummernkreis nicht möglich.")

        with tabs[4]:
            st.subheader("Bank-Reconciliation")
            if bank_file:
                df_bank = pd.read_csv(bank_file, sep=None, engine='python')
                st.success("Bankdaten erfolgreich geladen.")
                st.dataframe(df_bank.head(15), width='stretch')
            else:
                st.info("Bitte Bank-CSV laden, um den Abgleich zu nutzen.")
    else:
        st.info("👋 Willkommen! Bitte Datei laden, Mapping prüfen und Analyse starten.")
