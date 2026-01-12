import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO, StringIO
from datetime import datetime
import numpy as np
import re

# --- 1. DESIGN & KONFIGURATION ---
st.set_page_config(page_title="Sohn Consult Executive BI", page_icon="👔", layout="wide")

# CSS für professionellen Look
st.markdown(
    """
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
    </style>
    """,
    unsafe_allow_html=True,
)

# --- 2. HILFSFUNKTIONEN (STABILITY LAYER) ---
def clean_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """
    Bereinigt den DataFrame beim Import, um Abstürze durch
    gemischte Spaltennamen oder leere Bereiche zu verhindern.
    """
    # Spaltennamen zu Strings erzwingen und bereinigen
    df.columns = [str(c).strip() for c in df.columns]

    # Leere Spalten entfernen
    cols_to_keep = [c for c in df.columns if "Unnamed" not in c and c not in ("nan", "", "None")]
    df = df[cols_to_keep].copy()

    # Komplett leere Zeilen löschen
    df.dropna(how="all", inplace=True)
    return df


def format_euro(val) -> str:
    if pd.isna(val) or val is None:
        return "0,00 €"
    return f"{val:,.2f} €".replace(",", "X").replace(".", ",").replace("X", ".")


def to_excel(df: pd.DataFrame) -> bytes:
    output = BytesIO()
    # openpyxl ist i.d.R. stabiler verfügbar als xlsxwriter
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Analyse")
    return output.getvalue()


def find_idx(cols, keys) -> int:
    for i, c in enumerate(cols):
        if any(k.lower() in str(c).lower() for k in keys):
            return i
    return 0


def parse_money_series(s: pd.Series) -> pd.Series:
    # robustes Parsing für typische DE Formate (1.234,56) sowie bereits numerische Werte
    if pd.api.types.is_numeric_dtype(s):
        return pd.to_numeric(s, errors="coerce")
    x = s.astype(str).str.strip()
    x = x.str.replace("€", "", regex=False).str.replace(" ", "", regex=False)
    x = x.str.replace(".", "", regex=False).str.replace(",", ".", regex=False)
    return pd.to_numeric(x, errors="coerce")


# --- 3. APP HEADER ---
st.title("👔 Sohn Consult Strategic BI Dashboard")
st.caption("Version 2026.9 Stable Core Forensic & Cashflow")
st.markdown("---")

# --- 4. DATA IMPORT ---
col_u1, col_u2 = st.columns(2)
with col_u1:
    fibu_file = st.file_uploader("📂 1. Fibu Datei laden (XLSX/CSV)", type=["xlsx", "csv"])
with col_u2:
    bank_file = st.file_uploader("🏦 2. OPTIONAL Bank CSV laden", type=["csv"])

df_work = None
start_btn = False
date_range = None
sel_kunden = None

if fibu_file:
    with st.sidebar:
        st.header("⚙️ Konfiguration")
        mode = st.radio("Format", ["Standard Excel CSV", "DATEV Export"])
        header_row = st.number_input("Header Zeile", min_value=1, value=3)

        try:
            # Einlesen
            if mode == "DATEV Export":
                content = fibu_file.getvalue().decode("latin-1", errors="ignore")
                df_raw = pd.read_csv(StringIO(content), sep=None, engine="python", skiprows=1)
            else:
                if fibu_file.name.lower().endswith(".csv"):
                    df_raw = pd.read_csv(fibu_file, sep=None, engine="python")
                else:
                    df_raw = pd.read_excel(fibu_file, header=int(header_row - 1))

            # Bereinigung
            df_work = clean_dataframe(df_raw)
            cols = df_work.columns.tolist()

            st.subheader("📍 Mapping")

            c_dat = st.selectbox("Rechnungsdatum", cols, index=find_idx(cols, ["datum", "belegdat"]))
            c_fae = st.selectbox("Fälligkeit", cols, index=find_idx(cols, ["fällig", "faellig", "termin"]))
            c_nr = st.selectbox("RE Nummer", cols, index=find_idx(cols, ["nummer", "belegfeld", "re", "rechnung"]))
            c_kun = st.selectbox("Kunde", cols, index=find_idx(cols, ["kunde", "name", "debitor"]))
            c_bet = st.selectbox("Betrag", cols, index=find_idx(cols, ["brutto", "betrag", "umsatz", "summe"]))
            c_pay = st.selectbox("Zahldatum", cols, index=find_idx(cols, ["gezahlt", "ausgleich", "eingang", "zahlung"]))

            # Typ Handling
            df_work[c_dat] = pd.to_datetime(df_work[c_dat], errors="coerce")

            # Text Kopie für Anzeige
            df_work["Fällig_Text"] = df_work[c_fae].astype(str).fillna("")

            # Echtes Datum für Rechnen
            df_work[c_fae] = pd.to_datetime(df_work[c_fae], errors="coerce")
            df_work[c_pay] = pd.to_datetime(df_work[c_pay], errors="coerce")

            # Betrag robust parsen
            df_work[c_bet] = parse_money_series(df_work[c_bet])

            # Mindestvalidierung
            df_work = df_work.dropna(subset=[c_dat, c_bet]).copy()

            # FILTER
            st.markdown("### 🔍 Filter")

            if c_kun in df_work.columns:
                k_list = sorted(df_work[c_kun].dropna().astype(str).unique().tolist())
                sel_kunden = st.multiselect("Kunden", options=k_list, default=k_list)
            else:
                sel_kunden = []

            if not df_work.empty:
                min_d = df_work[c_dat].min().date()
                max_d = df_work[c_dat].max().date()
                date_range = st.date_input("Zeitraum", [min_d, max_d])
                start_btn = st.button("🚀 ANALYSE STARTEN", use_container_width=True)
            else:
                st.error("Keine gültigen Daten nach Import und Bereinigung.")
                start_btn = False

        except Exception as e:
            st.error(f"Fehler beim Import oder Mapping: {e}")
            start_btn = False

# --- 5. ANALYSE ---
if df_work is not None and start_btn and date_range and len(date_range) == 2:
    # Robuster Kunden Filter: wenn leer, dann kein Filter
    kunden_mask = df_work[c_kun].isin(sel_kunden) if sel_kunden else True

    mask = (
        (df_work[c_dat].dt.date >= date_range[0]) &
        (df_work[c_dat].dt.date <= date_range[1]) &
        kunden_mask
    )
    f_df = df_work.loc[mask].copy()

    today = pd.Timestamp(datetime.now().date())
    df_offen = f_df[f_df[c_pay].isna()].copy()
    df_paid = f_df[~f_df[c_pay].isna()].copy()

    tabs = st.tabs(["📊 Performance", "🔴 Forderungen", "💎 Strategie", "🔍 Forensik", "🏦 Bank"])

    # TAB 1: PERFORMANCE
    with tabs[0]:
        k1, k2, k3, k4 = st.columns(4)
        rev = float(f_df[c_bet].sum()) if not f_df.empty else 0.0

        k1.metric("Gesamtumsatz", format_euro(rev))
        k2.metric("Offene Posten", format_euro(float(df_offen[c_bet].sum()) if not df_offen.empty else 0.0))

        dso = (df_paid[c_pay] - df_paid[c_dat]).dt.days.mean() if not df_paid.empty else np.nan
        k3.metric("Ø Zahlungsdauer", f"{dso:.1f} Tage" if pd.notna(dso) and dso > 0 else "N/A")
        k4.metric("Belege", int(len(f_df)))

        c1, c2 = st.columns([2, 1])

        with c1:
            if not f_df.empty:
                f_df["Monat"] = f_df[c_dat].dt.strftime("%Y-%m")
                mon_chart = f_df.groupby("Monat", as_index=False)[c_bet].sum()
                fig_bar = px.bar(mon_chart, x="Monat", y=c_bet, title="Umsatz")
                st.plotly_chart(fig_bar, use_container_width=True)
            else:
                st.info("Keine Daten im gewählten Zeitraum.")

        with c2:
            if not f_df.empty:
                f_sorted = f_df.sort_values(c_dat).copy()
                f_sorted["Kumuliert"] = f_sorted[c_bet].cumsum()
                fig_area = px.area(f_sorted, x=c_dat, y="Kumuliert", title="Wachstum")
                st.plotly_chart(fig_area, use_container_width=True)
            else:
                st.info("Keine Daten für Wachstumskurve.")

    # TAB 2: FORDERUNGEN
    with tabs[1]:
        st.subheader("Forderungs Management")

        c_op1, c_op2 = st.columns([1, 2])

        # Verzug berechnen, NaT bleibt NaN
        if not df_offen.empty:
            df_offen["Verzug"] = np.where(
                df_offen[c_fae].isna(),
                np.nan,
                (today - df_offen[c_fae]).dt.days
            )
        else:
            df_offen["Verzug"] = pd.Series(dtype="float")

        def get_bucket(d):
            if pd.isna(d):
                return "Unbekannt"
            if d <= 0:
                return "1. Pünktlich"
            if d <= 30:
                return "2. 1-30 Tage"
            if d <= 60:
                return "3. 31-60 Tage"
            return "4. > 60 Tage"

        with c_op1:
            if not df_offen.empty:
                df_offen["Bucket"] = df_offen["Verzug"].apply(get_bucket)
                pie_data = df_offen.groupby("Bucket", as_index=False)[c_bet].sum()
                fig_pie = px.pie(pie_data, values=c_bet, names="Bucket", hole=0.5, title="Risiko")
                st.plotly_chart(fig_pie, use_container_width=True)
            else:
                st.info("Keine offenen Posten im gewählten Zeitraum.")

        with c_op2:
            if not df_offen.empty:
                # Prognose nur auf echte Fälligkeitsdaten
                df_predict = (
                    df_offen.dropna(subset=[c_fae])
                    .groupby(c_fae, as_index=False)[c_bet]
                    .sum()
                )
                if not df_predict.empty:
                    df_predict["Size_Safe"] = df_predict[c_bet].abs().clip(lower=0.1)
                    fig_scat = px.scatter(
                        df_predict,
                        x=c_fae,
                        y=c_bet,
                        size="Size_Safe",
                        title="Cash Inflow Prognose"
                    )
                    st.plotly_chart(fig_scat, use_container_width=True)
                else:
                    st.info("Keine offenen Posten mit Fälligkeit für Prognose.")
            else:
                st.info("Keine offenen Posten für Prognose.")

        # Tabelle mit sicherem Text Datum
        if not df_offen.empty:
            try:
                show_cols = [c_dat, "Fällig_Text", c_kun, c_bet, "Verzug"]
                st.dataframe(
                    df_offen.sort_values("Verzug", ascending=False)[show_cols],
                    column_config={c_bet: st.column_config.NumberColumn(format="%.2f €")},
                    use_container_width=True
                )
            except Exception:
                st.dataframe(df_offen, use_container_width=True)

            st.download_button(
                "📥 Excel OP Liste",
                data=to_excel(df_offen),
                file_name="OP_Liste.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        else:
            st.info("Keine Daten für OP Tabelle oder Export.")

    # TAB 3: STRATEGIE
    with tabs[2]:
        st.subheader("ABC Analyse")

        if not f_df.empty and c_kun in f_df.columns:
            abc = (
                f_df.groupby(c_kun, as_index=False)[c_bet]
                .sum()
                .sort_values(c_bet, ascending=False)
            )
            fig_abc = px.bar(abc.head(15), x=c_kun, y=c_bet, title="Top Kunden")
            st.plotly_chart(fig_abc, use_container_width=True)

            top3_share = (abc[c_bet].head(3).sum() / rev * 100) if rev > 0 else 0
            st.metric("Klumpenrisiko Top 3", f"{top3_share:.1f}%")
        else:
            st.info("Nicht genug Daten für ABC Analyse.")

    # TAB 4: FORENSIK
    with tabs[3]:
        st.subheader("🔍 Forensik")

        l1, l2 = st.columns(2)

        with l1:
            st.markdown("**Logik Check**")
            if not f_df.empty:
                err = f_df[(~f_df[c_pay].isna()) & (f_df[c_pay] < f_df[c_dat])]
                if not err.empty:
                    st.error(f"Fehler: {len(err)} Zahlung vor Rechnung")
                    st.dataframe(err, use_container_width=True)
                else:
                    st.success("Logik OK")
            else:
                st.info("Keine Daten für Logik Check.")

        with l2:
            st.markdown("**Nummernkreis**")
            if not f_df.empty and c_nr in f_df.columns:
                try:
                    def get_n(x):
                        found = re.findall(r"\d+", str(x))
                        return int(found[-1]) if found else None

                    nums = pd.Series(f_df[c_nr].apply(get_n)).dropna().astype(int)
                    nums = np.array(sorted(nums.unique()))
                    if len(nums) > 1:
                        full = np.arange(nums.min(), nums.max() + 1)
                        miss = np.setdiff1d(full, nums)
                        if len(miss) > 0:
                            st.warning(f"Nummern fehlen: {len(miss)}")
                            st.write(miss[:20])
                        else:
                            st.success("Lückenlos")
                    else:
                        st.info("Nicht genug Nummern für Prüfung.")
                except Exception:
                    st.info("Nummernkreis nicht prüfbar.")
            else:
                st.info("Keine Daten für Nummernkreis Prüfung.")

    # TAB 5: BANK
    with tabs[4]:
        st.subheader("Bank Abgleich")
        if bank_file:
            try:
                df_bank = pd.read_csv(bank_file, sep=None, engine="python")
                st.success("Bankdaten geladen.")
                st.dataframe(df_bank.head(50), use_container_width=True)
            except Exception as e:
                st.error(f"Fehler beim Lesen der Bank CSV: {e}")
        else:
            st.info("Bitte Bank CSV laden.")

else:
    st.info("👋 Bitte Datei laden und starten.")
