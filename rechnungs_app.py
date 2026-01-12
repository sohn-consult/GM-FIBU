import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO, StringIO
from datetime import datetime
import numpy as np

# --- 1. CONFIG & STYLING (Modern Business) ---
st.set_page_config(page_title="Sohn-Consult | Executive BI", page_icon="📈", layout="wide")

st.markdown("""
    <style>
    .stApp { background-color: #F8FAFC; }
    [data-testid="stSidebar"] { background-color: #FFFFFF; border-right: 1px solid #E2E8F0; }
    h1, h2, h3 { color: #1E3A8A; font-family: 'Inter', sans-serif; font-weight: 800; }
    .stMetric { background-color: white; padding: 20px; border-radius: 15px; box-shadow: 0 10px 15px -3px rgba(0,0,0,0.1); border-left: 8px solid #1E3A8A; }
    .stTabs [aria-selected="true"] { background-color: #1E3A8A !important; color: white !important; border-radius: 8px; }
    .status-box { padding: 15px; border-radius: 10px; margin-bottom: 10px; font-weight: 600; }
    </style>
    """, unsafe_allow_html=True)

def format_euro(val):
    if pd.isna(val): return "0,00 €"
    return f"{val:,.2f} €".replace(",", "X").replace(".", ",").replace("X", ".")

# --- 2. HEADER & LOGIK ---
st.title("👔 Sohn-Consult | Executive Strategy Dashboard")
st.caption("Advanced Financial Intelligence & Receivables Management")
st.markdown("---")

uploaded_file = st.file_uploader("📂 Fibu-Daten (XLSX/CSV) für Analyse hochladen", type=["xlsx", "csv"])

if uploaded_file:
    with st.sidebar:
        st.header("⚙️ Analyse-Parameter")
        mode = st.radio("Daten-Typ", ["Standard Excel/CSV", "DATEV Export"])
        header_row = st.number_input("Header-Zeile", min_value=1, value=3)
        
        try:
            if mode == "DATEV Export":
                content = uploaded_file.getvalue().decode('latin-1', errors='ignore')
                df_raw = pd.read_csv(StringIO(content), sep=None, engine='python', skiprows=1)
            else:
                if uploaded_file.name.endswith('.csv'):
                    df_raw = pd.read_csv(uploaded_file, sep=None, engine='python')
                else:
                    df_raw = pd.read_excel(uploaded_file, header=header_row-1)
            
            # Cleaning Column Names
            df_raw.columns = [str(c).strip() for c in df_raw.columns]
            cols = [c for c in df_raw.columns if "Unnamed" not in c]
            df = df_raw[cols].dropna(how='all', axis=0)

            st.subheader("📍 Spalten-Mapping")
            def find_c(keys):
                for i, c in enumerate(cols):
                    if any(k.lower() in c.lower() for k in keys): return i
                return 0

            c_dat = st.selectbox("Rechnungsdatum", cols, index=find_c(["datum", "belegdat"]))
            c_fael = st.selectbox("Fälligkeitsdatum", cols, index=find_c(["fällig", "termin"]))
            c_kun = st.selectbox("Kunde", cols, index=find_c(["kunde", "name"]))
            c_bet = st.selectbox("Betrag", cols, index=find_c(["betrag", "brutto", "umsatz"]))
            c_pay = st.selectbox("Zahldatum", cols, index=find_c(["gezahlt", "ausgleich", "eingang"]))
            
            # Transformation
            df[c_dat] = pd.to_datetime(df[c_dat], errors='coerce')
            df[c_fael] = pd.to_datetime(df[c_fael], errors='coerce')
            df[c_pay] = pd.to_datetime(df[c_pay], errors='coerce')
            if df[c_bet].dtype == 'object':
                df[c_bet] = pd.to_numeric(df[c_bet].astype(str).str.replace('.', '').str.replace(',', '.'), errors='coerce')
            df = df.dropna(subset=[c_dat, c_bet])
            
            st.success("Daten verarbeitet!")
            start_btn = st.button("🚀 ANALYSE STARTEN", use_container_width=True)

        except Exception as e:
            st.error(f"Fehler: {e}")
            start_btn = False

    if start_btn:
        # --- CALCULATION ENGINE ---
        today = pd.Timestamp(datetime.now().date())
        offen_mask = df[c_pay].isna()
        df_offen = df[offen_mask].copy()
        df_paid = df[~offen_mask].copy()
        
        # 1. Aging Buckets
        df_offen['Verzug'] = (today - df_offen[c_fael]).dt.days
        def bucket(d):
            if d <= 0: return "Pünktlich/Offen"
            if d <= 30: return "1-30 Tage"
            if d <= 60: return "31-60 Tage"
            if d <= 90: return "61-90 Tage"
            return "> 90 Tage"
        df_offen['Bucket'] = df_offen['Verzug'].apply(bucket)

        # 2. Performance Metrics
        df['Monat'] = df[c_dat].dt.to_period('M').astype(str)
        monthly = df.groupby('Monat')[c_bet].sum().reset_index()
        monthly['Wachstum'] = monthly[c_bet].pct_change() * 100

        # --- TABS ---
        t1, t2, t3, t4 = st.tabs(["🚀 Strategic Summary", "📊 Performance Analysis", "🔴 Receivables Management", "🔍 Forensik"])

        with t1:
            st.subheader("Executive Highlights")
            c1, c2, c3, c4 = st.columns(4)
            rev = df[c_bet].sum()
            ope = df_offen[c_bet].sum()
            c1.metric("Gesamtumsatz", format_euro(rev))
            c2.metric("Offen (Risk)", format_euro(ope), delta=f"{(ope/rev*100):.1f}% Quote", delta_color="inverse")
            
            # DSO (Days Sales Outstanding) Approximation
            dso = (df_paid[c_pay] - df_paid[c_dat]).dt.days.mean()
            c3.metric("Ø Zahlungsziel (DSO)", f"{dso:.1f} Tage" if not pd.isna(dso) else "N/A")
            c4.metric("Ø Rechnungsgröße", format_euro(df[c_bet].mean()))

            st.markdown("---")
            cc1, cc2 = st.columns([2, 1])
            with cc1:
                # Kumulierter Umsatz S-Kurve
                df_sorted = df.sort_values(c_dat)
                df_sorted['Kumuliert'] = df_sorted[c_bet].cumsum()
                fig_s = px.area(df_sorted, x=c_dat, y='Kumuliert', title="Kumulierter Jahresumsatz (Wachstumspfad)", color_discrete_sequence=['#1E3A8A'])
                st.plotly_chart(fig_s, width='stretch')
            with cc2:
                st.info("**Berater-Hinweis:**")
                if ope/rev > 0.2:
                    st.warning("⚠️ Die Forderungsquote ist mit über 20% kritisch. Fokus auf Mahnwesen legen!")
                else:
                    st.success("✅ Die Liquiditätsbindung ist im grünen Bereich.")

        with t2:
            st.subheader("Umsatz- & Wachstumsdynamik")
            col_p1, col_p2 = st.columns(2)
            with col_p1:
                fig_bar = px.bar(monthly, x='Monat', y=c_bet, text_auto='.2s', title="Monatliche Performance (Netto)", color_discrete_sequence=['#3B82F6'])
                st.plotly_chart(fig_bar, width='stretch')
            with col_p2:
                fig_growth = px.line(monthly, x='Monat', y='Wachstum', markers=True, title="MoM Wachstumsrate (%)", color_discrete_sequence=['#EF4444'])
                st.plotly_chart(fig_growth, width='stretch')

        with t3:
            st.subheader("Aging Structure & Risk Profile")
            col_a1, col_a2 = st.columns([1, 2])
            with col_a1:
                # Aging Donut
                aging_sum = df_offen.groupby('Bucket')[c_bet].sum().reset_index()
                fig_age = px.pie(aging_sum, values=c_bet, names='Bucket', hole=0.5, title="Überfälligkeiten nach Clustern",
                                 color_discrete_map={"Pünktlich/Offen":"#10B981", "1-30 Tage":"#F59E0B", "31-60 Tage":"#F97316", "> 90 Tage":"#EF4444"})
                st.plotly_chart(fig_age, width='stretch')
            with col_a2:
                # Cashflow Timeline (Predictive)
                df_predict = df_offen.groupby(c_fael)[c_bet].sum().reset_index()
                fig_pre = px.scatter(df_predict, x=c_fael, y=c_bet, size=c_bet, title="Erwarteter Cash-Inflow (nach Fälligkeit)", color_discrete_sequence=['#10B981'])
                st.plotly_chart(fig_pre, width='stretch')
            
            st.markdown("### Top 10 Debitoren-Risiken")
            top_debt = df_offen.groupby(c_kun)[c_bet].sum().reset_index().sort_values(c_bet, ascending=False).head(10)
            st.table(top_debt.assign(Betrag=top_debt[c_bet].apply(format_euro)).drop(columns=[c_bet]))

        with t4:
            st.subheader("Audit & Forensic-Check")
            c_f1, c_f2 = st.columns(2)
            with c_f1:
                # Wochentags-Analyse (Unregelmäßigkeiten bei Buchungen am WE?)
                df['Wochentag'] = df[c_dat].dt.day_name()
                day_order = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
                fig_day = px.bar(df.groupby('Wochentag')[c_bet].count().reindex(day_order).reset_index(), x='Wochentag', y=c_bet, title="Belegfrequenz nach Wochentag")
                st.plotly_chart(fig_day, width='stretch')
            with c_f2:
                st.markdown("**Prüfungsergebnisse:**")
                logik = df[df[c_pay] < df[c_dat]]
                if not logik.empty:
                    st.error(f"❌ {len(logik)} Rechnungen haben ein Zahldatum VOR dem Belegdatum.")
                else:
                    st.success("✅ Zeitliche Logik der Zahlungen ist konsistent.")
                
                # Check für lückenlose Nummern (Top 5 Lücken)
                try:
                    nums = pd.to_numeric(df[c_nr], errors='coerce').dropna().astype(int).sort_values().unique()
                    diffs = np.diff(nums)
                    luecken = np.where(diffs > 1)[0]
                    if len(luecken) > 0:
                        st.warning(f"⚠️ {len(luecken)} Lücken im Nummernkreis detektiert.")
                    else:
                        st.success("✅ Rechnungsnummernkreis scheint lückenlos.")
                except:
                    st.info("Nummernkreis-Prüfung bei diesem Format übersprungen.")
