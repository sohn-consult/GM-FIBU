import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO, StringIO
from datetime import datetime
import numpy as np
import re

# --- 1. DESIGN & KONFIGURATION (Standard 2026) ---
st.set_page_config(page_title="Sohn-Consult | Strategic BI", page_icon="👔", layout="wide")

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

# --- 2. ZENTRALE REINIGUNGS-FUNKTION (IMPORT-FILTER) ---
def sanitize_data(df):
    """Bereinigt den DataFrame für maximale Stabilität."""
    # Spaltennamen zu Strings machen und säubern
    df.columns = [str(c).strip() for c in df.columns]
    # Entferne komplett leere Spalten oder 'Unnamed'
    cols_to_keep = [c for c in df.columns if "Unnamed" not in c and c.lower() != "nan" and c != ""]
    df = df[cols_to_keep].dropna(how='all', axis=0)
    return df

def format_euro(val):
    if pd.isna(val) or val is None: return "0,00 €"
    return f"{val:,.2f} €".replace(",", "X").replace(".", ",").replace("X", ".")

def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Analyse_SohnConsult')
    return output.getvalue()

st.title("👔 Sohn-Consult | Strategic BI Dashboard")
st.caption("Professionelle Fibu-Analyse & Forensic - Version 2026.5 (Stable)")
st.markdown("---")

# --- 3. MULTI-UPLOAD ---
col_u1, col_u2 = st.columns(2)
with col_u1:
    fibu_file = st.file_uploader("📂 1. Fibu/Debitoren laden", type=["xlsx", "csv"])
with col_u2:
    bank_file = st.file_uploader("🏦 2. OPTIONAL: Bankumsätze (CSV) laden", type=["csv"])

if fibu_file:
    with st.sidebar:
        st.header("⚙️ Parameter")
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
            
            # --- DATEN-REINIGUNG BEIM IMPORT ---
            df_work = sanitize_data(df_raw)
            cols = df_work.columns.tolist()

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

            # Transformationen
            df_work[c_dat] = pd.to_datetime(df_work[c_dat], errors='coerce')
            df_work['Fällig_Display'] = df_work[c_fae].astype(str).fillna("unbekannt")
            df_work[c_fae] = pd.to_datetime(df_work[c_fae], errors='coerce')
            df_work[c_pay] = pd.to_datetime(df_work[c_pay], errors='coerce')
            
            if df_work[c_bet].dtype == 'object':
                df_work[c_bet] = pd.to_numeric(df_work[c_bet].astype(str).str.replace('.', '').str.replace(',', '.'), errors='coerce')
            
            df_work = df_work.dropna(subset=[c_dat, c_bet])

            st.markdown("### 🔍 Filter")
            k_list = sorted(df_work[c_kun].dropna().unique().tolist())
            sel_kunden = st.multiselect("Kunden", options=k_list, default=k_list)
            date_range = st.date_input("Zeitraum", [df_work[c_dat].min().date(), df_work[c_dat].max().date()])

            start_btn = st.button("🚀 ANALYSE STARTEN", width='stretch')

        except Exception as e:
            st.error(f"Fehler beim Daten-Import: {e}")
            start_btn = False

    # --- 4. HAUPTANALYSE ---
    if start_btn and len(date_range) == 2:
        mask = (df_work[c_dat].dt.date >= date_range[0]) & (df_work[c_dat].dt.date <= date_range[1]) & (df_work[c_kun].isin(sel_kunden))
        f_df = df_work[mask].copy()

        today = pd.Timestamp(datetime.now().date())
        df_offen = f_df[f_df[c_pay].isna()].copy()
        df_paid = f_df[~f_df[c_pay].isna()].copy()

        tabs = st.tabs(["📊 Performance", "🔴 Aging & OP", "💎 Strategie", "🔍 Forensik", "🏦 Bank-Match"])

        with tabs[0]:
            k1, k2, k3, k4 = st.columns(4)
            rev = f_df[c_bet].sum()
            k1.metric("Gesamtumsatz", format_euro(rev))
            k2.metric("Offene Posten", format_euro(df_offen[c_bet].sum()))
            dso = (df_paid[c_pay] - df_paid[c_dat]).dt.days.mean() if not df_paid.empty else 0
            k3.metric("Ø Zahlungsdauer", f"{dso:.1f} Tage" if dso > 0 else "N/A")
            k4.metric("Anzahl Belege", len(f_df))

            c_p1, c_p2 = st.columns([2, 1])
            with c_p1:
                f_df['Monat'] = f_df[c_dat].dt.strftime('%Y-%m')
                st.plotly_chart(px.bar(f_df.groupby('Monat')[c_bet].sum().reset_index(), x='Monat', y=c_bet, color_discrete_sequence=['#1E3A8A'], title="Umsatzverlauf"), width='stretch')
            with c_p2:
                f_df = f_df.sort_values(c_dat)
                f_df['Kumuliert'] = f_df[c_bet].cumsum()
                st.plotly_chart(px.area(f_df, x=c_dat, y='Kumuliert', title="Jahres-Wachstumspfad", color_discrete_sequence=['#3B82F6']), width='stretch')

        with tabs[1]:
            st.subheader("Forderungs-Management & Aging")
            col_a1, col_a2 = st.columns([1, 2])
            with col_a1:
                df_offen['Verzug'] = (today - df_offen[c_fae]).dt.days
                def bucket(d):
                    if pd.isna(d): return "5. Unbekannt"
                    if d <= 0: return "1. Pünktlich"
                    if d <= 30: return "2. 1-30 Tage"
                    if d <= 60: return "3. 31-60 Tage"
                    return "4. > 60 Tage"
                df_offen['Bucket'] = df_offen['Verzug'].apply(bucket)
                st.plotly_chart(px.pie(df_offen.groupby('Bucket')[c_bet].sum().reset_index(), values=c_bet, names='Bucket', hole=0.5, title="Überfälligkeiten"), width='stretch')
            
            with col_a2:
                # --- FIX: SCATTER PLOT CRASH (Größe muss positiv sein) ---
                df_predict = df_offen.groupby(c_fae)[c_bet].sum().reset_index()
                if not df_predict.empty:
                    df_predict['Betrag_Abs'] = df_predict[c_bet].abs().clip(lower=0.1)
                    st.plotly_chart(px.scatter(df_predict, x=c_fae, y=c_bet, size='Betrag_Abs', title="Cashflow-Timeline", color_discrete_sequence=['#10B981']), width='stretch')
            
            # --- FIX: ARROW DISPLAY CRASH ---
            disp_df = df_offen[[c_dat, c_fae, c_kun, c_bet, 'Verzug']].copy()
            disp_df[c_dat] = disp_df[c_dat].dt.strftime('%d.%m.%Y')
            disp_df[c_fae] = df_offen['Fällig_Display']
            st.dataframe(disp_df.sort_values('Verzug', ascending=False), column_config={c_bet: st.column_config.NumberColumn(format="%.2f €")}, width='stretch')

        with tabs[2]:
            st.subheader("ABC-Analyse")
            abc = f_df.groupby(c_kun)[c_bet].sum().reset_index().sort_values(by=c_bet, ascending=False)
            st.plotly_chart(px.bar(abc.head(15), x=c_kun, y=c_bet, title="Top-Kunden-Ranking", color_discrete_sequence=['#1E3A8A']), width='stretch')
            top3 = (abc[c_bet].head(3).sum() / rev * 100) if rev > 0 else 0
            st.metric("Klumpenrisiko (Top 3)", f"{top3:.1f}%")

        with tabs[3]:
            st.subheader("🔍 Daten-Forensik")
            l1, l2 = st.columns(2)
            with l1:
                st.markdown("### Logik-Check")
                err = f_df[f_df[c_pay] < f_df[c_dat]]
                if not err.empty: st.error(f"❌ {len(err)} Fehler: Zahlung VOR Rechnung."); st.dataframe(err[[c_dat, c_pay, c_kun, c_bet]])
                else: st.success("✅ Datum-Logik konsistent.")
            with l2:
                st.markdown("### RE-Nummernkreis")
                try:
                    def ext_n(v):
                        n = re.findall(r'\d+', str(v))
                        return int(n[-1]) if n else None
                    nums = f_df[c_nr].apply(ext_n).dropna().sort_values().unique()
                    if len(nums) > 1:
                        miss = np.setdiff1d(np.arange(nums.min(), nums.max() + 1), nums)
                        if len(miss) > 0: st.warning(f"⚠️ {len(miss)} Nummern fehlen."); st.write(miss[:50])
                        else: st.success("✅ Nummernkreis lückenlos.")
                except: st.info("Check nicht möglich.")

        with tabs[4]:
            st.subheader("Bank-Matching")
            if bank_file:
                df_bank = pd.read_csv(bank_file, sep=None, engine='python')
                st.success("Bankdaten erfolgreich geladen.")
                st.dataframe(df_bank.head(15), width='stretch')
            else: st.info("Bank-CSV laden für automatischen Abgleich.")
    else:
        st.info("👋 Willkommen! Bitte Datei laden und 'Analyse starten' klicken.")
