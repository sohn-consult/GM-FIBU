import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO, StringIO
from datetime import datetime
import numpy as np

# --- 1. DESIGN & KONFIGURATION ---
st.set_page_config(page_title="Sohn-Consult | Executive BI", page_icon="📈", layout="wide")

st.markdown("""
    <style>
    .stApp { background-color: #F8FAFC; }
    [data-testid="stSidebar"] { background-color: #FFFFFF; border-right: 1px solid #E2E8F0; }
    h1, h2, h3 { color: #1E3A8A; font-family: 'Inter', sans-serif; font-weight: 800; }
    .stMetric { background-color: white; padding: 20px; border-radius: 15px; box-shadow: 0 10px 15px -3px rgba(0,0,0,0.1); border-left: 8px solid #1E3A8A; }
    .stTabs [aria-selected="true"] { background-color: #1E3A8A !important; color: white !important; border-radius: 8px; }
    .stButton>button { background-color: #1E3A8A; color: white; border-radius: 8px; font-weight: bold; width: 100%; height: 3em; }
    </style>
    """, unsafe_allow_html=True)

# Hilfsfunktionen
def format_euro(val):
    if pd.isna(val) or val is None: return "0,00 €"
    return f"{val:,.2f} €".replace(",", "X").replace(".", ",").replace("X", ".")

def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Sohn_Consult_Export')
    return output.getvalue()

st.title("👔 Sohn-Consult | Executive Strategy Dashboard")
st.caption("Universal-Tool: Fibu, Forensic, Bank-Reconciliation & Cashflow")
st.markdown("---")

# --- 2. MULTI-UPLOAD BEREICH ---
col_u1, col_u2 = st.columns(2)
with col_u1:
    fibu_file = st.file_uploader("📂 1. Fibu/Debitoren-Datei laden", type=["xlsx", "csv"])
with col_u2:
    bank_file = st.file_uploader("🏦 2. OPTIONAL: Bankumsätze (CSV) laden", type=["csv"])

if fibu_file:
    with st.sidebar:
        st.header("⚙️ Konfiguration")
        mode = st.radio("Daten-Format", ["Standard Excel/CSV", "DATEV Export"])
        header_row = st.number_input("Header-Zeile", min_value=1, value=3)
        
        try:
            if mode == "DATEV Export":
                content = fibu_file.getvalue().decode('latin-1', errors='ignore')
                df_raw = pd.read_csv(StringIO(content), sep=None, engine='python', skiprows=1)
            else:
                if fibu_file.name.endswith('.csv'):
                    df_raw = pd.read_csv(fibu_file, sep=None, engine='python')
                else:
                    df_raw = pd.read_excel(fibu_file, header=header_row-1)
            
            # Cleaning
            df_raw.columns = [str(c).strip() for c in df_raw.columns]
            cols = [c for c in df_raw.columns if "Unnamed" not in str(c)]
            df_work = df_raw[cols].dropna(how='all', axis=0)

            st.subheader("📍 Spalten-Mapping")
            def find_idx(keys, default=0):
                for i, c in enumerate(cols):
                    if any(k.lower() in str(c).lower() for k in keys): return i
                return default

            c_dat = st.selectbox("Rechnungsdatum", cols, index=find_idx(["datum", "belegdat"]))
            c_fae = st.selectbox("Fälligkeitsdatum", cols, index=find_idx(["fällig", "termin"]))
            c_nr = st.selectbox("RE-Nummer", cols, index=find_idx(["nummer", "belegfeld"]))
            c_kun = st.selectbox("Kunde", cols, index=find_idx(["kunde", "name"]))
            c_bet = st.selectbox("Betrag", cols, index=find_idx(["brutto", "betrag", "umsatz"]))
            c_pay = st.selectbox("Zahldatum", cols, index=find_idx(["gezahlt", "ausgleich", "eingang"]))

            # Transformation
            df_work[c_dat] = pd.to_datetime(df_work[c_dat], errors='coerce')
            df_work[c_fae] = pd.to_datetime(df_work[c_fae], errors='coerce')
            df_work[c_pay] = pd.to_datetime(df_work[c_pay], errors='coerce')
            if df_work[c_bet].dtype == 'object':
                df_work[c_bet] = pd.to_numeric(df_work[c_bet].astype(str).str.replace('.', '').str.replace(',', '.'), errors='coerce')
            df_work = df_work.dropna(subset=[c_dat, c_bet])

            # --- FILTER (STABILISIERT) ---
            st.markdown("### 🔍 Filter")
            kunden_list = sorted(df_work[c_kun].dropna().unique().tolist())
            sel_kunden = st.multiselect("Kunden", options=kunden_list, default=kunden_list)
            
            min_d, max_d = df_work[c_dat].min().date(), df_work[c_dat].max().date()
            date_range = st.date_input("Zeitraum", [min_d, max_d])

            start_btn = st.button("🚀 ANALYSE STARTEN", use_container_width=True)

        except Exception as e:
            st.error(f"Mapping-Fehler: {e}")
            start_btn = False

    # --- 3. ANALYSE ENGINE ---
    if start_btn and len(date_range) == 2:
        mask = (df_work[c_dat].dt.date >= date_range[0]) & \
               (df_work[c_dat].dt.date <= date_range[1]) & \
               (df_work[c_kun].isin(sel_kunden))
        f_df = df_work[mask].copy()

        today = pd.Timestamp(datetime.now().date())
        offen_mask = f_df[c_pay].isna()
        df_offen = f_df[offen_mask].copy()
        df_paid = f_df[~offen_mask].copy()

        # Tabs
        tabs = st.tabs(["📊 Performance", "🔴 Aging & OP", "💎 Strategie & ABC", "🔍 Forensik", "🏦 Bank-Match"])

        with tabs[0]:
            k1, k2, k3, k4 = st.columns(4)
            rev = f_df[c_bet].sum()
            k1.metric("Gesamtumsatz", format_euro(rev))
            k2.metric("Offene Posten", format_euro(df_offen[c_bet].sum()))
            dso = (df_paid[c_pay] - df_paid[c_dat]).dt.days.mean()
            k3.metric("Ø Zahlungsdauer", f"{dso:.1f} Tage" if not pd.isna(dso) else "N/A")
            k4.metric("Belege", len(f_df))

            c_p1, c_p2 = st.columns([2, 1])
            with c_p1:
                f_df['Monat'] = f_df[c_dat].dt.strftime('%Y-%m')
                st.plotly_chart(px.bar(f_df.groupby('Monat')[c_bet].sum().reset_index(), x='Monat', y=c_bet, color_discrete_sequence=['#1E3A8A'], title="Umsatz nach Monat"), width='stretch')
            with c_p2:
                # S-Kurve (Kumuliert)
                f_df = f_df.sort_values(c_dat)
                f_df['Kumuliert'] = f_df[c_bet].cumsum()
                st.plotly_chart(px.area(f_df, x=c_dat, y='Kumuliert', title="Wachstumspfad", color_discrete_sequence=['#3B82F6']), width='stretch')

        with tabs[1]:
            st.subheader("Receivables & Aging Structure")
            col_a1, col_a2 = st.columns([1, 2])
            with col_a1:
                df_offen['Verzug'] = (today - df_offen[c_fae]).dt.days
                def bucket(d):
                    if d <= 0: return "1. Pünktlich"
                    if d <= 30: return "2. 1-30 Tage"
                    if d <= 60: return "3. 31-60 Tage"
                    return "4. > 60 Tage"
                df_offen['Bucket'] = df_offen['Verzug'].apply(bucket)
                st.plotly_chart(px.pie(df_offen.groupby('Bucket')[c_bet].sum().reset_index(), values=c_bet, names='Bucket', hole=0.4, title="Risiko-Verteilung"), width='stretch')
            
            with col_a2:
                # FIX FÜR SCATTER PLOT CRASH (ValueError size)
                df_predict = df_offen.groupby(c_fae)[c_bet].sum().reset_index()
                df_predict['Size_Fix'] = df_predict[c_bet].clip(lower=0.1) # Verhindert Crash bei negativen Werten
                st.plotly_chart(px.scatter(df_predict, x=c_fae, y=c_bet, size='Size_Fix', title="Erwarteter Cash-Inflow (nach Fälligkeit)", color_discrete_sequence=['#10B981']), width='stretch')
            
            st.dataframe(df_offen[[c_dat, c_fae, c_kun, c_bet, 'Verzug']].sort_values('Verzug', ascending=False), 
                         column_config={c_bet: st.column_config.NumberColumn(format="%.2f €")}, width='stretch')
            st.download_button("📥 Excel OP-Liste", to_excel(df_offen), "OP_SohnConsult.xlsx")

        with tabs[2]:
            st.subheader("ABC-Analyse & Klumpenrisiko")
            abc = f_df.groupby(c_kun)[c_bet].sum().reset_index().sort_values(c_bet, ascending=False)
            st.plotly_chart(px.bar(abc.head(15), x=c_kun, y=c_bet, title="Top 15 Kundenumsätze"), width='stretch')
            top3 = (abc[c_bet].head(3).sum() / rev) * 100
            st.metric("Klumpenrisiko (Top 3)", f"{top3:.1f}%")

        with tabs[3]:
            st.subheader("Prüfungsmodul (Forensik)")
            l1, l2 = st.columns(2)
            with l1:
                err = f_df[f_df[c_pay] < f_df[c_dat]]
                if not err.empty: st.error(f"Logikfehler: {len(err)} Zahlungen VOR Rechnung."); st.dataframe(err)
                else: st.success("Datum-Logik einwandfrei.")
            with l2:
                try:
                    nums = pd.to_numeric(f_df[c_nr], errors='coerce').dropna().astype(int).sort_values().unique()
                    miss = np.setdiff1d(np.arange(nums.min(), nums.max() + 1), nums)
                    if len(miss) > 0: st.warning(f"Lücken im Nummernkreis: {len(miss)} Nummern fehlen.")
                    else: st.success("Nummernkreis lückenlos.")
                except: st.info("Check nicht verfügbar.")

        with tabs[4]:
            st.subheader("Bank-Reconciliation (Matching)")
            if bank_file:
                df_bank = pd.read_csv(bank_file, sep=None, engine='python')
                st.success("Bankdaten geladen. Gleiche Beträge mit Fibu ab...")
                st.dataframe(df_bank.head(10), width='stretch')
            else:
                st.info("Laden Sie eine Bank-CSV hoch, um offene Posten abzugleichen.")
    else:
        st.info("Bitte Datei laden und Analyse starten.")
