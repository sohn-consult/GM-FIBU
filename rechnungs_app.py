import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO, StringIO
from datetime import datetime
import numpy as np
import re
import warnings

# -----------------------------
# 1) PAGE SETUP + DESIGN
# -----------------------------
st.set_page_config(page_title="Sohn Consult Executive BI", page_icon="👔", layout="wide")

st.markdown(
    """
    <style>
    .stApp { background-color: #F8FAFC; }
    [data-testid="stSidebar"] { background-color: #F1F5F9; border-right: 1px solid #CBD5E1; }
    h1, h2, h3 { color: #1E3A8A; font-family: Inter, sans-serif; font-weight: 700; }
    .stMetric {
        background-color: #FFFFFF; padding: 18px; border-radius: 12px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.05); border-top: 5px solid #1E3A8A;
    }
    .stButton>button {
        background-color: #1E3A8A; color: white; border-radius: 8px;
        font-weight: 700; width: 100%; height: 3.2em;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("👔 Sohn Consult Strategic BI Dashboard")
st.caption("Stable Core 2026: Forensic & Cashflow")
st.markdown("---")

# -----------------------------
# 2) HELPERS
# -----------------------------
def clean_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    df.columns = [str(c).strip() for c in df.columns]
    keep = [c for c in df.columns if "Unnamed" not in c and c not in ("", "nan", "None")]
    df = df[keep].copy()
    df.dropna(how="all", inplace=True)
    return df

def parse_money(s: pd.Series) -> pd.Series:
    if pd.api.types.is_numeric_dtype(s):
        return pd.to_numeric(s, errors="coerce")
    x = s.astype(str).str.strip()
    x = x.str.replace("€", "", regex=False).str.replace(" ", "", regex=False)
    x = x.str.replace(".", "", regex=False).str.replace(",", ".", regex=False)
    return pd.to_numeric(x, errors="coerce")

def parse_date(s: pd.Series) -> pd.Series:
    """
    Warnungsfreies Datum Parsing:
    - niemals "infer format" auslösen
    - explizite Formate für die häufigsten Muster
    - exotische Fälle: silent fallback (Warnungen werden geschluckt)
    """
    if pd.api.types.is_datetime64_any_dtype(s):
        return pd.to_datetime(s, errors="coerce")

    x = s.astype(str).str.strip()
    x = x.replace({"": np.nan, "nan": np.nan, "None": np.nan})

    # Muster
    iso_dt = x.str.match(r"^\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2}$", na=False)
    iso_d  = x.str.match(r"^\d{4}-\d{2}-\d{2}$", na=False)
    dot_d  = x.str.match(r"^\d{1,2}\.\d{1,2}\.\d{2,4}$", na=False)
    slash_d = x.str.match(r"^\d{1,2}/\d{1,2}/\d{2,4}$", na=False)
    dash_d = x.str.match(r"^\d{1,2}-\d{1,2}-\d{2,4}$", na=False)

    out = pd.Series(pd.NaT, index=x.index, dtype="datetime64[ns]")

    # ISO datetime
    if iso_dt.any():
        out.loc[iso_dt] = pd.to_datetime(x.loc[iso_dt], format="%Y-%m-%d %H:%M:%S", errors="coerce")

    # ISO date
    if iso_d.any():
        out.loc[iso_d] = pd.to_datetime(x.loc[iso_d], format="%Y-%m-%d", errors="coerce")

    # DE dot date: 12.01.2026
    if dot_d.any():
        # akzeptiert auch 2-stellige Jahre schlecht, aber coerced
        out.loc[dot_d] = pd.to_datetime(x.loc[dot_d], format="%d.%m.%Y", errors="coerce")

    # Slash date: 12/01/2026 -> hier unterstellen wir dayfirst (DE Logik)
    if slash_d.any():
        out.loc[slash_d] = pd.to_datetime(x.loc[slash_d], format="%d/%m/%Y", errors="coerce")

    # Dash date: 12-01-2026 -> dayfirst (DE Logik)
    if dash_d.any():
        out.loc[dash_d] = pd.to_datetime(x.loc[dash_d], format="%d-%m-%Y", errors="coerce")

    # Rest: letzter Fallback, aber Warnungen unterdrücken (wichtig, wenn Warnungen als Fehler laufen)
    rest = out.isna() & x.notna()
    if rest.any():
        with warnings.catch_warnings():
            warnings.simplefilter("ignore", category=UserWarning)
            out.loc[rest] = pd.to_datetime(x.loc[rest], errors="coerce")

    return out

def format_euro(val) -> str:
    if pd.isna(val) or val is None:
        return "0,00 €"
    return f"{float(val):,.2f} €".replace(",", "X").replace(".", ",").replace("X", ".")

def to_excel_bytes(df: pd.DataFrame) -> bytes:
    out = BytesIO()
    with pd.ExcelWriter(out) as writer:
        df.to_excel(writer, index=False, sheet_name="Analyse")
    return out.getvalue()

def find_idx(cols, keys) -> int:
    for i, c in enumerate(cols):
        c_low = str(c).lower()
        if any(k in c_low for k in keys):
            return i
    return 0

def make_arrow_safe(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    for c in out.columns:
        if out[c].dtype == "object":
            out[c] = out[c].astype("string")
    return out

# -----------------------------
# 3) UPLOADS
# -----------------------------
col1, col2 = st.columns(2)
with col1:
    fibu_file = st.file_uploader("📂 1. Fibu Datei laden (XLSX/CSV)", type=["xlsx", "csv"])
with col2:
    bank_file = st.file_uploader("🏦 2. OPTIONAL Bank CSV laden", type=["csv"])

if not fibu_file:
    st.info("👋 Bitte Datei laden und starten.")
    st.stop()

# -----------------------------
# 4) IMPORT
# -----------------------------
with st.sidebar:
    st.header("⚙️ Konfiguration")
    mode = st.radio("Format", ["Standard Excel CSV", "DATEV Export"])
    header_row = st.number_input("Header Zeile", min_value=1, value=3)

try:
    if mode == "DATEV Export":
        content = fibu_file.getvalue().decode("latin-1", errors="ignore")
        df_raw = pd.read_csv(StringIO(content), sep=None, engine="python", skiprows=1)
    else:
        if fibu_file.name.lower().endswith(".csv"):
            df_raw = pd.read_csv(fibu_file, sep=None, engine="python")
        else:
            df_raw = pd.read_excel(fibu_file, header=int(header_row - 1))

    df = clean_dataframe(df_raw)
    cols = df.columns.tolist()
except Exception as e:
    st.error("Import fehlgeschlagen.")
    st.exception(e)
    st.stop()

# -----------------------------
# 5) MAPPING + NORMALISIERUNG
# -----------------------------
with st.sidebar:
    st.subheader("📍 Mapping")
    c_dat = st.selectbox("Rechnungsdatum", cols, index=find_idx(cols, ["datum", "belegdat"]))
    c_fae = st.selectbox("Fälligkeit", cols, index=find_idx(cols, ["fällig", "faellig", "termin"]))
    c_nr  = st.selectbox("RE Nummer", cols, index=find_idx(cols, ["nummer", "belegfeld", "rechnung", "beleg", "re-nr", "re nr"]))
    c_kun = st.selectbox("Kunde", cols, index=find_idx(cols, ["kunde", "name", "debitor"]))
    c_bet = st.selectbox("Betrag", cols, index=find_idx(cols, ["brutto", "betrag", "umsatz", "summe"]))
    c_pay = st.selectbox("Zahldatum", cols, index=find_idx(cols, ["gezahlt", "ausgleich", "eingang", "zahlung"]))

df[c_dat] = parse_date(df[c_dat])
df[c_fae] = parse_date(df[c_fae])
df[c_pay] = parse_date(df[c_pay])
df[c_bet] = parse_money(df[c_bet])

df["Fällig_Text"] = df[c_fae].astype("string").fillna("")

df = df.dropna(subset=[c_dat, c_bet]).copy()
if df.empty:
    st.error("Nach Bereinigung keine gültigen Datensätze übrig.")
    st.stop()

# Arrow kritische Spalten fix typisieren
if c_nr in df.columns:
    df[c_nr] = df[c_nr].astype("string")
if c_kun in df.columns:
    df[c_kun] = df[c_kun].astype("string")

# -----------------------------
# 6) FILTER
# -----------------------------
with st.sidebar:
    st.markdown("### 🔍 Filter")

    if c_kun in df.columns:
        kunden = sorted(df[c_kun].dropna().astype(str).unique().tolist())
        sel_kunden = st.multiselect("Kunden", options=kunden, default=kunden)
    else:
        sel_kunden = []

    min_d = df[c_dat].min().date()
    max_d = df[c_dat].max().date()
    date_range = st.date_input("Zeitraum", [min_d, max_d])

    start_btn = st.button("🚀 ANALYSE STARTEN", width="stretch")

if not start_btn:
    st.info("Konfiguration prüfen und Analyse starten.")
    st.stop()

if not date_range or len(date_range) != 2:
    st.error("Bitte einen Zeitraum mit Start und Ende auswählen.")
    st.stop()

kunden_mask = df[c_kun].isin(sel_kunden) if sel_kunden else True
mask = (
    (df[c_dat].dt.date >= date_range[0]) &
    (df[c_dat].dt.date <= date_range[1]) &
    kunden_mask
)

f_df = df.loc[mask].copy()
if f_df.empty:
    st.warning("Keine Datensätze im gewählten Filter.")
    st.stop()

today = pd.Timestamp(datetime.now().date())
df_offen = f_df[f_df[c_pay].isna()].copy()
df_paid = f_df[~f_df[c_pay].isna()].copy()

tabs = st.tabs(["📊 Performance", "🔴 Forderungen", "💎 Strategie", "🔍 Forensik", "🏦 Bank"])

# -----------------------------
# TAB 1: PERFORMANCE
# -----------------------------
with tabs[0]:
    k1, k2, k3, k4 = st.columns(4)
    rev = float(f_df[c_bet].sum())
    open_sum = float(df_offen[c_bet].sum()) if not df_offen.empty else 0.0

    k1.metric("Gesamtumsatz", format_euro(rev))
    k2.metric("Offene Posten", format_euro(open_sum))

    dso = (df_paid[c_pay] - df_paid[c_dat]).dt.days.mean() if not df_paid.empty else np.nan
    k3.metric("Ø Zahlungsdauer", f"{dso:.1f} Tage" if pd.notna(dso) and dso > 0 else "N/A")
    k4.metric("Belege", int(len(f_df)))

    cA, cB = st.columns([2, 1])

    with cA:
        f_df["Monat"] = f_df[c_dat].dt.strftime("%Y-%m")
        mon = f_df.groupby("Monat", as_index=False)[c_bet].sum()
        st.plotly_chart(px.bar(mon, x="Monat", y=c_bet, title="Umsatz"), width="stretch")

    with cB:
        f_sorted = f_df.sort_values(c_dat).copy()
        f_sorted["Kumuliert"] = f_sorted[c_bet].cumsum()
        st.plotly_chart(px.area(f_sorted, x=c_dat, y="Kumuliert", title="Wachstum"), width="stretch")

# -----------------------------
# TAB 2: FORDERUNGEN
# -----------------------------
with tabs[1]:
    st.subheader("Forderungs Management")

    if df_offen.empty:
        st.info("Keine offenen Posten im Filter.")
    else:
        df_offen["Verzug"] = np.where(
            df_offen[c_fae].isna(),
            np.nan,
            (today - df_offen[c_fae]).dt.days
        )

        def bucket(d):
            if pd.isna(d): return "Unbekannt"
            if d <= 0: return "1. Pünktlich"
            if d <= 30: return "2. 1-30 Tage"
            if d <= 60: return "3. 31-60 Tage"
            return "4. > 60 Tage"

        c1, c2 = st.columns([1, 2])

        with c1:
            df_offen["Bucket"] = df_offen["Verzug"].apply(bucket)
            pie = df_offen.groupby("Bucket", as_index=False)[c_bet].sum()
            st.plotly_chart(px.pie(pie, values=c_bet, names="Bucket", hole=0.5, title="Risiko"), width="stretch")

        with c2:
            df_pred = (
                df_offen.dropna(subset=[c_fae])
                .groupby(c_fae, as_index=False)[c_bet]
                .sum()
            )
            if df_pred.empty:
                st.info("Keine offenen Posten mit Fälligkeit für Prognose.")
            else:
                df_pred["Size"] = df_pred[c_bet].abs().clip(lower=0.1)
                st.plotly_chart(
                    px.scatter(df_pred, x=c_fae, y=c_bet, size="Size", title="Cash Inflow Prognose"),
                    width="stretch"
                )

        show_cols = [c_dat, "Fällig_Text", c_kun, c_nr, c_bet, "Verzug"]
        show_cols = [c for c in show_cols if c in df_offen.columns]

        view = df_offen.sort_values("Verzug", ascending=False)[show_cols]
        st.dataframe(
            make_arrow_safe(view),
            column_config={c_bet: st.column_config.NumberColumn(format="%.2f €")},
            width="stretch"
        )

        st.download_button(
            "📥 Excel OP Liste",
            data=to_excel_bytes(df_offen),
            file_name="OP_Liste.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width="stretch"
        )

# -----------------------------
# TAB 3: STRATEGIE
# -----------------------------
with tabs[2]:
    st.subheader("ABC Analyse")
    if c_kun not in f_df.columns:
        st.info("Keine Kunden Spalte verfügbar.")
    else:
        abc = f_df.groupby(c_kun, as_index=False)[c_bet].sum().sort_values(c_bet, ascending=False)
        st.plotly_chart(px.bar(abc.head(15), x=c_kun, y=c_bet, title="Top Kunden"), width="stretch")
        top3_share = (abc[c_bet].head(3).sum() / rev * 100) if rev > 0 else 0
        st.metric("Klumpenrisiko Top 3", f"{top3_share:.1f}%")

# -----------------------------
# TAB 4: FORENSIK
# -----------------------------
with tabs[3]:
    st.subheader("🔍 Forensik")
    l1, l2 = st.columns(2)

    with l1:
        st.markdown("**Logik Check**")
        err = f_df[(~f_df[c_pay].isna()) & (f_df[c_pay] < f_df[c_dat])]
        if err.empty:
            st.success("Logik OK")
        else:
            st.error(f"Fehler: {len(err)} Zahlung vor Rechnung")
            st.dataframe(make_arrow_safe(err), width="stretch")

    with l2:
        st.markdown("**Nummernkreis**")
        if c_nr not in f_df.columns:
            st.info("Keine RE Nummer Spalte verfügbar.")
        else:
            try:
                def get_n(x):
                    found = re.findall(r"\d+", str(x))
                    return int(found[-1]) if found else None

                nums = pd.Series(f_df[c_nr].apply(get_n)).dropna().astype(int)
                nums = np.array(sorted(nums.unique()))
                if len(nums) <= 1:
                    st.info("Nicht genug Nummern für Prüfung.")
                else:
                    full = np.arange(nums.min(), nums.max() + 1)
                    miss = np.setdiff1d(full, nums)
                    if len(miss) == 0:
                        st.success("Lückenlos")
                    else:
                        st.warning(f"Nummern fehlen: {len(miss)}")
                        st.write(miss[:20])
            except Exception:
                st.info("Nummernkreis nicht prüfbar.")

# -----------------------------
# TAB 5: BANK
# -----------------------------
with tabs[4]:
    st.subheader("Bank Abgleich")
    if not bank_file:
        st.info("Bitte Bank CSV laden.")
    else:
        try:
            df_bank = pd.read_csv(bank_file, sep=None, engine="python")
            st.success("Bankdaten geladen.")
            st.dataframe(make_arrow_safe(df_bank.head(100)), width="stretch")
        except Exception as e:
            st.error("Fehler beim Lesen der Bank CSV")
            st.exception(e)
