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
    [data-testid="stSidebar"] .stMarkdown p, [data-testid="stSidebar"] label { 
        color: #0F172A !important; font-weight: 600 !important; 
    }
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

# Hilfsfunktionen
def format_euro(val):
    if pd.isna(val) or val is None: return "0,00 €"
    return f"{val:,.2f} €".replace(",", "X").replace(".", ",").replace("X", ".")

def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Sohn_Consult_Export')
    return output.getvalue()

st.title("👔 Sohn-Consult | Strategic BI Dashboard")
st.caption("Forensic Edition: Advanced Data Integrity & Business Intelligence")
st.markdown("---")

# --- 2. MULTI-UPLOAD BEREICH ---
col_u1, col_u2 = st.columns(2)
with col_u1:
    fibu_file = st.file_uploader("📂 1. Fibu/Debitoren laden", type=["xlsx", "csv"])
with col_u2:
    bank_file = st.file_uploader("🏦 2. OPTIONAL: Bankumsätze (CSV) laden", type=["csv"])

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
                    df_raw = pd.read_excel(fibu_file, header=header_row-1)
            
            df_raw.columns = [str(c).strip() for c in df_raw.columns]
            cols = [c for c in df_raw.columns if "Unnamed" not in c and c != "nan"]
            df_work = df_raw[cols].dropna(how='all', axis=0).copy()

            st.subheader("📍 Spalten-Mapping")
            def find_idx(keys, default=0):
                for i, c in enumerate(cols):
                    if any(k.lower() in c.lower() for k in keys): return i
                return default

            c_dat = st.selectbox("Rechnungsdatum", cols, index=find_idx(["datum", "belegdat"]))
            c_fae = st.selectbox("Fälligkeitsdatum", cols, index=find_idx(["fällig", "termin"]))
            c_nr = st.selectbox("RE-Nummer", cols, index=find_idx(["nummer", "belegfeld"]))
            c_kun = st.selectbox("Kunde", cols, index=find_idx(["kunde", "name"]))
            c_bet = st.selectbox("Betrag", cols, index=find_idx(["brutto", "betrag", "umsatz"]))
            c_pay = st.selectbox("Zahldatum", cols, index=find_idx(["gezahlt", "ausgleich", "eingang"]))

            # Transformation
            df_work[c_dat] = pd.to_datetime(df_work[c_dat], errors='coerce')
            df_work['Fällig_Display'] = df_work[c_fae].astype(str) 
            df_work[c_fae] = pd.to_datetime(df_work[c_fae], errors='coerce')
            df_work[c_pay] = pd.to_datetime(df_work[c_pay], errors='coerce')
            
            if df_work[c_bet].dtype == 'object':
                df_work[c_bet] = pd.to_numeric(df_work[c_bet].astype(str).str.replace('.', '').str.replace(',', '.'), errors='coerce')
            
            df_work = df_work.dropna(subset=[c_dat, c_bet])

            st.markdown("### 🔍 Filter")
            kunden_list = sorted(df_work[c_kun].dropna().unique().tolist())
            sel_kunden = st.multiselect("Kunden", options=kunden_list, default=kunden_list)
            
            min_d, max_d = df_work[c_dat].min().date(), df_work[c_dat].max().date()
            date_range = st.date_input("Zeitraum", [min_d, max_d])

            start_btn = st.button("🚀 ANALYSE STARTEN", use_container_width=True)

        except Exception as e:
            st.error(f"Fehler beim Laden: {e}")
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

        tabs = st.tabs(["📊 Performance", "🔴 Aging & Offene Posten", "💎 Strategie & Risiko", "🔍 Forensik", "🏦 Bank-Match"])

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
                st.plotly_chart(px.bar(f_df.groupby('Monat')[c_bet].sum().reset_index(), x='Monat', y=c_bet, color_discrete_sequence=['#1E3A8A'], title="Umsatz pro Monat"), use_container_width=True)
            with c_p2:
                f_df = f_df.sort_values(c_dat)
                f_df['Kumuliert'] = f_df[c_bet].cumsum()
                st.plotly_chart(px.area(f_df, x=c_dat, y='Kumuliert', title="Wachstumspfad", color_discrete_sequence=['#3B82F6']), use_container_width=True)

        with tabs[1]:
            st.subheader("Forderungs-Management")
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
                st.plotly_chart(px.pie(df_offen.groupby('Bucket')[c_bet].sum().reset_index(), values=c_bet, names='Bucket', hole=0.5, title="Überfälligkeiten"), use_container_width=True)
            
            with col_a2:
                df_predict = df_offen.groupby(c_fae)[c_bet].sum().reset_index()
                if not df_predict.empty:
                    df_predict['Betrag_Abs'] = df_predict[c_bet].abs().clip(lower=0.1) 
                    st.plotly_chart(px.scatter(df_predict, x=c_fae, y=c_bet, size='Betrag_Abs', title="Cashflow-Prognose", color_discrete_sequence=['#10B981']), use_container_width=True)
            
            disp_df = df_offen[[c_dat, c_fae, c_kun, c_bet, 'Verzug']].copy()
            disp_df[c_dat] = disp_df[c_dat].dt.strftime('%d.%m.%Y')
            disp_df[c_fae] = disp_df[c_fae].dt.strftime('%d.%m.%Y').fillna("unbekannt")
            st.dataframe(disp_df.sort_values('Verzug', ascending=False), column_config={c_bet: st.column_config.NumberColumn(format="%.2f €")}, use_container_width=True)

        with tabs[2]:
            st.subheader("ABC-Analyse & Klumpenrisiko")
            abc = f_df.groupby(c_kun)[c_bet].sum().reset_index().sort_values(by=c_bet, ascending=False)
            st.plotly_chart(px.bar(abc.head(15), x=c_kun, y=c_bet, title="Top 15 Kundenumsätze", color_discrete_sequence=['#1E3A8A']), use_container_width=True)
            top3 = (abc[c_bet].head(3).sum() / rev) * 100 if rev > 0 else 0
            st.metric("Klumpenrisiko (Top 3 Kunden)", f"{top3:.1f}%")

        with tabs[3]:
            st.subheader("🔍 Forensik & Daten-Integrität")
            
            # 1. Daten-Vollständigkeit
            st.markdown("### 1. Datenqualitäts-Check")
            missing_count = f_df.isna().sum().sum()
            if missing_count == 0:
                st.success("✅ Datensatz ist vollständig (keine leeren Felder in den Hauptspalten).")
            else:
                st.info(f"ℹ️ Es wurden {missing_count} leere Felder im Datensatz gefunden.")

            l1, l2 = st.columns(2)
            with l1:
                st.markdown("### 2. Zeitliche Plausibilität")
                err = f_df[f_df[c_pay] < f_df[c_dat]]
                if not err.empty:
                    st.error(f"❌ {len(err)} Buchungsfehler gefunden: Zahlungseingang liegt zeitlich VOR dem Rechnungsdatum.")
                    st.dataframe(err[[c_dat, c_pay, c_kun, c_bet]])
                else:
                    st.success("✅ Zeitliche Logik einwandfrei (Zahldaten liegen nach Rechnungsdaten).")

            with l2:
                st.markdown("### 3. Nummernkreis-Analyse")
                try:
                    # Robuster RE-Nummern Check: Extrahiert nur Zahlen aus dem Text
                    def extract_number(val):
                        nums = re.findall(r'\d+', str(val))
                        return int(nums[-1]) if nums else None
                    
                    nums_extracted = f_df[c_nr].apply(extract_number).dropna().sort_values().unique()
                    
                    if len(nums_extracted) > 1:
                        miss = np.setdiff1d(np.arange(nums_extracted.min(), nums_extracted.max() + 1), nums_extracted)
                        if len(miss) > 0:
                            st.warning(f"⚠️ Es fehlen {len(miss)} Rechnungsnummern in der Abfolge.")
                            with st.expander("Fehlende Nummern anzeigen (Auszug)"):
                                st.write(miss[:50])
                        else:
                            st.success("✅ Rechnungsnummernkreis ist lückenlos.")
                    else:
                        st.info("ℹ️ Zu wenig Daten für eine Nummernkreis-Analyse.")
                except Exception as e:
                    st.info(f"ℹ️ Nummernkreis-Check konnte nicht ausgeführt werden.")

            st.markdown("### 4. Dubletten-Prüfung")
            dubs = f_df[f_df.duplicated(subset=[c_nr, c_bet], keep=False)]
            if not dubs.empty:
                st.warning(f"⚠️ {len(dubs)} mögliche Dubletten gefunden (identische Rechnungsnummer & Betrag).")
                st.dataframe(dubs[[c_dat, c_nr, c_kun, c_bet]])
            else:
                st.success("✅ Keine identischen Rechnungs-Dubletten gefunden.")

        with tabs[4]:
            st.subheader("Bank-Reconciliation (Matching)")
            if bank_file:
                df_bank = pd.read_csv(bank_file, sep=None, engine='python')
                st.success("Bankdaten geladen.")
                st.dataframe(df_bank.head(15), use_container_width=True)
            else:
                st.info("Laden Sie eine Bank-CSV hoch, um den Abgleich zu nutzen.")
    else:
        st.info("Warten auf Datei-Upload und Klick auf 'Analyse starten'.")
