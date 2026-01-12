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

# --- 2. HILFSFUNKTIONEN (STABILITY LAYER) ---
def clean_dataframe(df):
    """Reinigt den DataFrame, um Abstürze durch korrupte Header oder leere Spalten zu verhindern."""
    # 1. Alle Spaltennamen in Strings umwandeln und bereinigen
    df.columns = [str(c).strip() for c in df.columns]
    # 2. Leere Spalten oder 'Unnamed' entfernen
    cols_to_keep = [c for c in df.columns if "Unnamed" not in c and c != "nan" and c != ""]
    df = df[cols_to_keep].copy()
    # 3. Komplett leere Zeilen löschen
    df.dropna(how='all', inplace=True)
    return df

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
st.caption("Stabilisierte Version: Forensic, Reporting & Cashflow (2026.7)")
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
            # Einlesen
            if mode == "DATEV Export":
                # Latin-1 ist oft bei DATEV nötig
                content = fibu_file.getvalue().decode('latin-1', errors='ignore')
                df_raw = pd.read_csv(StringIO(content), sep=None, engine='python', skiprows=1)
            else:
                if fibu_file.name.endswith('.csv'):
                    df_raw = pd.read_csv(fibu_file, sep=None, engine='python')
                else:
                    df_raw = pd.read_excel(fibu_file, header=int(header_row-1))
            
            # --- CRITICAL FIX: DATEN BEREINIGEN ---
            df_work = clean_dataframe(df_raw)
            cols = df_work.columns.tolist()

            st.subheader("📍 Mapping")
            def find_idx(keys):
                for i, c in enumerate(cols):
                    if any(k.lower() in str(c).lower() for k in keys): return i
                return 0

            # Spaltenwahl
            c_dat = st.selectbox("Rechnungsdatum", cols, index=find_idx(["datum", "belegdat"]))
            c_fae = st.selectbox("Fälligkeitsdatum", cols, index=find_idx(["fällig", "termin"]))
            c_nr = st.selectbox("RE-Nummer", cols, index=find_idx(["nummer", "belegfeld"]))
            c_kun = st.selectbox("Kunde", cols, index=find_idx(["kunde", "name"]))
            c_bet = st.selectbox("Betrag", cols, index=find_idx(["brutto", "betrag", "umsatz"]))
            c_pay = st.selectbox("Zahldatum", cols, index=find_idx(["gezahlt", "ausgleich", "eingang"]))

            # Transformationen (Fehlertolerant)
            df_work[c_dat] = pd.to_datetime(df_work[c_dat], errors='coerce')
            
            # ARROW FIX: Wir speichern eine reine Text-Kopie für die Anzeige in Tabellen
            df_work['Fällig_Text'] = df_work[c_fae].astype(str).fillna("")
            
            # Echte Datums-Konvertierung für Berechnungen
            df_work[c_fae] = pd.to_datetime(df_work[c_fae], errors='coerce')
            df_work[c_pay] = pd.to_datetime(df_work[c_pay], errors='coerce')
            
            # Zahlen reinigen (Tausenderpunkte etc.)
            if df_work[c_bet].dtype == 'object':
                df_work[c_bet] = pd.to_numeric(
                    df_work[c_bet].astype(str).str.replace('.', '', regex=False).str.replace(',', '.', regex=False),
                    errors='coerce'
                )
            
            # Leere Datensätze entfernen
            df_work = df_work.dropna(subset=[c_dat, c_bet])

            st.markdown("### 🔍 Filter")
            # Sicherstellen, dass Kundenliste existiert
            if c_kun in df_work.columns:
                k_list = sorted(df_work[c_kun].dropna().unique().astype(str).tolist())
                sel_kunden = st.multiselect("Kunden filtern", options=k_list, default=k_list)
            else:
                sel_kunden = []
            
            if not df_work.empty:
                min_d, max_d = df_work[c_dat].min().date(), df_work[c_dat].max().date()
                date_range = st.date_input("Zeitraum", [min_d, max_d])
                start_btn = st.button("🚀 ANALYSE STARTEN", width='stretch') # Syntax 2026 konform
            else:
                st.error("Keine gültigen Daten nach Bereinigung erkannt.")
                start_btn = False

        except Exception as e:
            st.error(f"Kritischer Fehler beim Laden: {e}")
            start_btn = False

    # --- 5. ANALYSE LOGIK ---
    if start_btn and len(date_range) == 2:
        # Filter anwenden
        mask = (df_work[c_dat].dt.date >= date_range[0]) & \
               (df_work[c_dat].dt.date <= date_range[1]) & \
               (df_work[c_kun].isin(sel_kunden))
        f_df = df_work[mask].copy()

        # Status bestimmen
        today = pd.Timestamp(datetime.now().date())
        df_offen = f_df[f_df[c_pay].isna()].copy()
        df_paid = f_df[~f_df[c_pay].isna()].copy()

        # Tabs aufbauen
        tabs = st.tabs(["📊 Performance", "🔴 Forderungen (OP)", "💎 Strategie (ABC)", "🔍 Forensik", "🏦 Bank-Abgleich"])

        with tabs[0]:
            k1, k2, k3, k4 = st.columns(4)
            rev = f_df[c_bet].sum()
            k1.metric("Gesamtumsatz", format_euro(rev))
            k2.metric("Offene Posten", format_euro(df_offen[c_bet].sum()))
            
            dso = (df_paid[c_pay] - df_paid[c_dat]).dt.days.mean() if not df_paid.empty else 0
            k3.metric("Ø Zahlungsdauer", f"{dso:.1f} Tage" if dso > 0 else "N/A")
            k4.metric("Anzahl Belege", len(f_df))

            c1, c2 = st.columns([2, 1])
            with c1:
                f_df['Monat'] = f_df[c_dat].dt.strftime('%Y-%m')
                mon_chart = f_df.groupby('Monat')[c_bet].sum().reset_index()
                st.plotly_chart(px.bar(mon_chart, x='Monat', y=c_bet, title="Umsatzverlauf", color_discrete_sequence=['#1E3A8A']), width='stretch')
            with c2:
                f_df = f_df.sort_values(c_dat)
                f_df['Kumuliert'] = f_df[c_bet].cumsum()
                st.plotly_chart(px.area(f_df, x=c_dat, y='Kumuliert', title="Wachstumspfad", color_discrete_sequence=['#3B82F6']), width='stretch')

        with tabs[1]:
            st.subheader("Forderungs-Management & Aging")
            c_op1, c_op2 = st.columns([1, 2])
            
            # Verzug berechnen
            df_offen['Verzug'] = (today - df_offen[c_fae]).dt.days
            
            with c_op1:
                def get_bucket(d):
                    if pd.isna(d): return "Unbekannt"
                    if d <= 0: return "1. Pünktlich"
                    if d <= 30: return "2. 1-30 Tage"
                    if d <= 60: return "3. 31-60 Tage"
                    return "4. > 60 Tage"
                
                df_offen['Bucket'] = df_offen['Verzug'].apply(get_bucket)
                pie_data = df_offen.groupby('Bucket')[c_bet].sum().reset_index()
                st.plotly_chart(px.pie(pie_data, values=c_bet, names='Bucket', hole=0.5, title="Risiko-Verteilung"), width='stretch')

            with c_op2:
                # CRASH FIX: Scatter
