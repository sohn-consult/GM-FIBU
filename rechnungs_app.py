import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO, StringIO
from datetime import datetime
import numpy as np
import re

# --- 1. DESIGN & KONFIGURATION ---
st.set_page_config(page_title="Sohn-Consult | Executive BI", page_icon="👔", layout="wide")

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

# --- 2. HILFSFUNKTIONEN ---
def clean_dataframe(df):
    # Bereinigt Spaltennamen und entfernt leere Spalten
    df.columns = [str(c).strip() for c in df.columns]
    cols_to_keep = [c for c in df.columns if "Unnamed" not in c and c.lower() != "nan" and c != ""]
    return df[cols_to_keep].dropna(how='all', axis=0)

def format_euro(val):
    if pd.isna(val) or val is None: return "0,00 €"
    return f"{val:,.2f} €".replace(",", "X").replace(".", ",").replace("X", ".")

def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Analyse_Export')
    return output.getvalue()

# --- 3. APP HEADER ---
st.title("👔 Sohn-Consult | Strategic BI Dashboard")
st.caption("Stabilisierte Version 2026.8 - Ready for Consulting")
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
            
            # Datenbereinigung
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

            # Datums-Konvertierung
            df_work[c_dat] = pd.to_datetime(df_work[c_dat], errors='coerce')
            
            # Speichere Fälligkeit als Text für die Anzeige (verhindert Absturz)
            df_work['Fällig_Text'] = df_work[c_fae].astype(str).fillna("")
            
            # Echte Konvertierung für Rechnen
            df_work[c_fae] = pd.to_datetime(df_work[c_fae], errors='coerce')
            df_work[c_pay] = pd.to_datetime(df_work[c_pay], errors='coerce')
            
            # Zahlen reinigen
            if df_work[c_bet].dtype == 'object':
                df_work[c_bet] = pd.to_numeric(
                    df_work[c_bet].astype(str).str.replace('.', '', regex=False).str.replace(',', '.', regex=False),
                    errors='coerce'
                )
            
            df_work = df_work.dropna(subset=[c_dat, c_bet])

            st.markdown("### 🔍 Filter")
            if c_kun in df_work.columns:
                k_list = sorted(df_work[c_kun].dropna().unique().astype(str).tolist())
                sel_kunden = st.multiselect("Kunden filtern", options=k_list, default=k_list)
            else:
                sel_kunden = []
            
            if not df_work.empty:
                min_d, max_d = df_work[c_dat].min().date(), df_work[c_dat].max().date()
                date_range = st.date_input("Zeitraum", [min_d, max_d])
                start_btn = st.button("🚀 ANALYSE STARTEN", use_container_width=True)
            else:
                st.error("Keine gültigen Daten erkannt.")
                start_btn = False

        except Exception as e:
            st.error(f"Fehler: {e}")
            start_btn = False

    # --- 5. ANALYSE LOGIK ---
    if start_btn and len(date_range) == 2:
        mask = (df_work[c_dat].dt.date >= date_range[0]) & \
               (df_work[c_dat].dt.date <= date_range[1]) & \
               (df_work[c_kun].isin(sel_kunden))
        f_df = df_work[mask].copy()

        today = pd.Timestamp(datetime.now().date())
        df_offen = f_df[f_df[c_pay].isna()].copy()
        df_paid = f_df[~f_df[c_pay].isna()].copy()

        tabs = st.tabs(["📊 Performance", "🔴 Aging & OP", "💎 Strategie", "🔍 Forensik", "🏦 Bank-Match"])

        # TAB 1: PERFORMANCE
        with tabs[0]:
            k1, k2, k3, k4 = st.columns(4)
            rev = f_df[c_bet].sum()
            k1.metric("Gesamtumsatz", format_euro(rev))
            k2.metric("Offene Posten", format_euro(df_offen[c_bet].sum()))
            
            if not df_paid.empty:
                dso = (df_paid[c_pay] - df_paid[c_dat]).dt.days.mean()
                k3.metric("Ø Zahlungsdauer", f"{dso:.1f} Tage")
            else:
                k3.metric("Ø Zahlungsdauer", "N/A")
            k4.metric("Belege", len(f_df))

            c1, c2 = st.columns([2, 1])
            with c1:
                f_df['Monat'] = f_df[c_dat].dt.strftime('%Y-%m')
                mon_chart = f_df.groupby('Monat')[c_bet].sum().reset_index()
                st.plotly_chart(px.bar(mon_chart, x='Monat', y=c_bet, title="Umsatzverlauf", color_discrete_sequence=['#1E3A8A']), use_container_width=True)
            with c2:
                f_df = f_df.sort_values(c_dat)
                f_df['Kumuliert'] = f_df[c_bet].cumsum()
                st.plotly_chart(px.area(f_df, x=c_dat, y='Kumuliert', title="Wachstumspfad", color_discrete_sequence=['#3B82F6']), use_container_width=True)

        # TAB 2: AGING & OP (Hier war der Fehler!)
        with tabs[1]:
            st.subheader("Forderungs-Management")
            col_a1, col_a2 = st.columns([1, 2])
            
            # Verzug berechnen
            df_offen['Verzug'] = (today - df_offen[c_fae]).dt.days
            
            with col_a1:
                def get_bucket(d):
                    if pd.isna(d): return "Unbekannt"
                    if d <= 0: return "1. Pünktlich"
                    if d <= 30: return "2. 1-30 Tage"
                    if d <= 60: return "3. 31-60 Tage"
                    return "4. > 60 Tage"
                
                df_offen['Bucket'] = df_offen['Verzug'].apply(get_bucket)
                pie_data = df_offen.groupby('Bucket')[c_bet].sum().reset_index()
                st.plotly_chart(px.pie(pie_data, values=c_bet, names='Bucket', hole=0.5, title="Risikoprofil"), use_container_width=True)

            with col_a2:
                # Scatter Plot (Größe positiv erzwingen)
                df_predict = df_offen.groupby(c_fae)[c_bet].sum().reset_index()
                if not df_predict.empty:
                    df_predict['Size_Safe'] = df_predict[c_bet].abs().clip(lower=0.1)
                    fig_scat = px.scatter(df_predict, x=c_fae, y=c_bet, size='Size_Safe', 
                                        title="Cash-Inflow Prognose", color_discrete_sequence=['#10B981'])
                    st.plotly_chart(fig_scat, use_container_width=True)
                else:
                    st.info("Keine Daten für Prognose.")
            
            # Tabelle mit Text-Spalte für Datum (verhindert Absturz)
            try:
                st.dataframe(
                    df_offen.sort_values('Verzug', ascending=False)[[c_dat, 'Fällig_Text', c_kun, c_bet, 'Verzug']],
                    column_config={c_bet: st.column_config.NumberColumn(format="%.2f €")},
                    use_container_width=True
                )
            except:
                st.dataframe(df_offen)
            
            st.download_button("📥 Excel-Liste", to_excel(df_offen), "OP_Liste.xlsx")

        # TAB 3: STRATEGIE
        with tabs[2]:
            st.subheader("ABC-Analyse")
            abc = f_df.groupby(c_kun)[c_bet].sum().reset_index().sort_values(c_bet, ascending=False)
            st.plotly_chart(px.bar(abc.head(15), x=c_kun, y=c_bet, title="Top-Kunden", color_discrete_sequence=['#1E3A8A']), use_container_width=True)
            
            top3_share = (abc[c_bet].head(3).sum() / rev * 100) if rev > 0 else 0
            st.metric("Klumpenrisiko (Top 3)", f"{top3_share:.1f}%")

        # TAB 4: FORENSIK
        with tabs[3]:
            st.subheader("🔍 Daten-Forensik")
            l1, l2 = st.columns(2)
            with l1:
                st.markdown("**Logik-Prüfung**")
                err_log = f_df[f_df[c_pay] < f_df[c_dat]]
                if not err_log.empty:
                    st.error(f"❌ {len(err_log)} Fehler: Zahlung vor Rechnung!")
                    st.dataframe(err_log)
                else:
                    st.success("✅ Zeitliche Logik korrekt.")
            
            with l2:
                st.markdown("**Nummernkreis**")
                try:
                    def get_num(x):
                        nums = re.findall(r'\d+', str(x))
                        return int(nums[-1]) if nums else None
                    
                    nr_series = f_df[c_nr].apply(get_num).dropna().unique()
                    if len(nr_series) > 1:
                        full = np.arange(nr_series.min(), nr_series.max() + 1)
                        missing = np.setdiff1d(full, nr_series)
                        if len(missing) > 0:
                            st.warning(f"⚠️ {len(missing)} Nummern fehlen.")
                            with st.expander("Details"): st.write(missing[:50])
                        else:
                            st.success("✅ Nummernkreis lückenlos.")
                except:
                    st.info("Nummernkreis nicht prüfbar.")

        # TAB 5: BANK
        with tabs[4]:
            st.subheader("Bank-Abgleich")
            if bank_file:
                df_bank = pd.read_csv(bank_file, sep=None, engine='python')
                st.success("Bankdaten geladen.")
                st.dataframe(df_bank.head(), use_container_width=True)
            else:
                st.info("Bitte Bank-CSV laden.")

    else:
        st.info("👋 Willkommen! Datei laden und 'Analyse starten' klicken.")
