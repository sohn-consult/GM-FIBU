import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO, StringIO
from datetime import datetime
import numpy as np
import re
import warnings
import traceback
import sys

warnings.filterwarnings("ignore", message="Could not infer format*", category=UserWarning)
warnings.filterwarnings("ignore", message="Parsing dates in %Y-%m-%d %H:%M:%S*", category=UserWarning)

# -----------------------------
# PAGE SETUP
# -----------------------------
st.set_page_config(page_title="Sohn Consult Executive BI", page_icon="👔", layout="wide")

st.markdown(
    """
    <style>
    .stApp { background-color: #F8FAFC; }
    [data-testid="stSidebar"] { background-color: #F1F5F9; border-right: 1px solid #CBD5E1; }
    h1, h2, h3 { color: #1E3A8A; font-family: Inter, sans-serif; font-weight: 700; }
    .stButton>button {
        background-color: #1E3A8A; color: white; border-radius: 8px;
        font-weight: 700; width: 100%; height: 3.2em;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

HEADER_KEYWORDS = [
    "kunde", "debitor", "re-nr", "re nr", "re-nr.", "rechnung", "beleg",
    "re-datum", "re datum", "datum", "fällig", "faellig", "termin",
    "betrag", "brutto", "netto", "summe", "monat", "gezahlt", "eingang",
    "zahlung", "ausgleich", "abzug"
]

def make_unique_columns(cols):
    seen = {}
    out = []
    for c in cols:
        base = str(c).strip()
        if base == "" or base.lower() in ("nan", "none"):
            base = "Spalte"
        if base not in seen:
            seen[base] = 1
            out.append(base)
        else:
            seen[base] += 1
            out.append(f"{base}_{seen[base]}")
    return out

def make_arrow_safe(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    for c in out.columns:
        if out[c].dtype == "object":
            out[c] = out[c].astype("string")
    return out

def parse_money(s: pd.Series) -> pd.Series:
    if pd.api.types.is_numeric_dtype(s):
        return pd.to_numeric(s, errors="coerce")
    x = s.astype(str).str.strip()
    x = x.str.replace("€", "", regex=False).str.replace(" ", "", regex=False)
    x = x.str.replace(".", "", regex=False).str.replace(",", ".", regex=False)
    return pd.to_numeric(x, errors="coerce")

def parse_date_strict(s: pd.Series) -> pd.Series:
    if pd.api.types.is_datetime64_any_dtype(s):
        return pd.to_datetime(s, errors="coerce")

    x = s.astype(str).str.strip()
    x = x.replace({"": np.nan, "nan": np.nan, "None": np.nan})

    out = pd.Series(pd.NaT, index=x.index, dtype="datetime64[ns]")

    iso_dt = x.str.match(r"^\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2}$", na=False)
    iso_d  = x.str.match(r"^\d{4}-\d{2}-\d{2}$", na=False)
    de_dot = x.str.match(r"^\d{1,2}\.\d{1,2}\.\d{2,4}$", na=False)

    if iso_dt.any():
        out.loc[iso_dt] = pd.to_datetime(x.loc[iso_dt], format="%Y-%m-%d %H:%M:%S", errors="coerce")
    if iso_d.any():
        out.loc[iso_d] = pd.to_datetime(x.loc[iso_d], format="%Y-%m-%d", errors="coerce")
    if de_dot.any():
        out.loc[de_dot] = pd.to_datetime(x.loc[de_dot], format="%d.%m.%Y", errors="coerce")

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
        cl = str(c).lower()
        if any(k in cl for k in keys):
            return i
    return 0

def clean_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    df.dropna(how="all", inplace=True)
    return df

def looks_like_header_cell(v: str) -> bool:
    if not v:
        return False
    vl = v.strip().lower()
    if vl in ("nan", "none"):
        return False
    return any(k in vl for k in HEADER_KEYWORDS)

def normalize_header_from_first_row(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df

    df2 = df.copy()
    row0 = df2.iloc[0].astype(str).str.strip()

    new_cols = []
    header_hits = 0

    for i in range(len(df2.columns)):
        v = row0.iloc[i] if i < len(row0) else ""
        if looks_like_header_cell(v):
            new_cols.append(v)
            header_hits += 1
        else:
            new_cols.append(f"col_{i:02d}")

    if header_hits >= 2:
        df2 = df2.iloc[1:].copy()
        df2.columns = make_unique_columns(new_cols)
        df2.dropna(how="all", inplace=True)
        if len(df2) > 0 and df2.iloc[0].isna().all():
            df2 = df2.iloc[1:].copy()
    else:
        df2.columns = make_unique_columns([str(c) for c in df2.columns])

    return df2

def crash_report(e: Exception, context: dict | None = None):
    st.error("App Fehler. Hier ist der technische Report. Damit können wir es final fixen.")
    st.write("Exception Typ:", type(e).__name__)
    st.write("Exception Text:", str(e))
    st.code(traceback.format_exc())

    if context:
        st.subheader("Debug Kontext")
        for k, v in context.items():
            try:
                st.write(k, v)
            except Exception:
                st.write(k, "nicht darstellbar")

def main():
    st.title("👔 Sohn Consult Strategic BI Dashboard")
    st.caption("Stable Core 2026: Forensic & Cashflow")
    st.markdown("---")

    col1, col2 = st.columns(2)
    with col1:
        fibu_file = st.file_uploader("📂 1. Fibu Datei laden (XLSX CSV)", type=["xlsx", "csv"])
    with col2:
        bank_file = st.file_uploader("🏦 2. OPTIONAL Bank CSV laden", type=["csv"])

    if not fibu_file:
        st.info("Bitte Datei laden.")
        return

    with st.sidebar:
        st.header("⚙️ Konfiguration")
        mode = st.radio("Format", ["Standard Excel CSV", "DATEV Export"])
        header_row = st.number_input("Header Zeile", min_value=1, value=3)
        show_diag = st.toggle("Diagnose Mode", value=True)

    # IMPORT
    try:
        if mode == "DATEV Export":
            content = fibu_file.getvalue().decode("latin-1", errors="ignore")
            df_raw = pd.read_csv(StringIO(content), sep=None, engine="python", skiprows=1)
            df = clean_dataframe(df_raw)
        else:
            if fibu_file.name.lower().endswith(".csv"):
                df_raw = pd.read_csv(fibu_file, sep=None, engine="python")
                df = clean_dataframe(df_raw)
            else:
                df_raw = pd.read_excel(fibu_file, header=int(header_row - 1))
                df = clean_dataframe(df_raw)
                df = normalize_header_from_first_row(df)

        if df.empty:
            st.error("Import Ergebnis ist leer.")
            return

        df.columns = make_unique_columns(df.columns.tolist())
        cols = df.columns.tolist()

    except Exception as e:
        crash_report(e)
        return

    if show_diag:
        with st.expander("🧪 Diagnose Import", expanded=True):
            st.write("Dateiname:", fibu_file.name)
            st.write("Shape:", df.shape)
            st.write("Spalten:", cols[:50])
            dtypes_view = pd.DataFrame({"Spalte": df.columns, "dtype": [str(df[c].dtype) for c in df.columns]})
            st.dataframe(make_arrow_safe(dtypes_view), width="stretch")
            st.dataframe(make_arrow_safe(df.head(20)), width="stretch")

    # MAPPING
    with st.sidebar:
        st.subheader("📍 Mapping")
        c_dat = st.selectbox("Rechnungsdatum", cols, index=find_idx(cols, ["re-datum", "re datum", "datum"]))
        c_fae = st.selectbox("Fälligkeit", cols, index=find_idx(cols, ["fällig", "faellig", "termin"]))
        c_nr  = st.selectbox("RE Nummer", cols, index=find_idx(cols, ["re-nr", "re nr", "nummer", "beleg"]))
        c_kun = st.selectbox("Kunde", cols, index=find_idx(cols, ["kunde", "name", "debitor"]))
        c_bet = st.selectbox("Betrag", cols, index=find_idx(cols, ["betrag (brutto)", "brutto", "betrag", "netto", "summe", "umsatz"]))
        c_pay = st.selectbox("Zahldatum", cols, index=find_idx(cols, ["gezahlt", "eingang", "zahlung", "ausgleich"]))

    # NORMALISIERUNG
    try:
        df["Fällig_Text"] = df[c_fae].astype("string").fillna("") if c_fae in df.columns else pd.Series([""] * len(df), dtype="string")

        df[c_dat] = parse_date_strict(df[c_dat])
        df[c_pay] = parse_date_strict(df[c_pay])
        df["_Fällig_Datum"] = parse_date_strict(df[c_fae]) if c_fae in df.columns else pd.Series(pd.NaT, index=df.index)

        df[c_bet] = parse_money(df[c_bet])

        if c_nr in df.columns:
            df[c_nr] = df[c_nr].astype("string")
        if c_kun in df.columns:
            df[c_kun] = df[c_kun].astype("string")

        df = df.dropna(subset=[c_dat, c_bet]).copy()
        if df.empty:
            st.error("Nach Bereinigung keine gültigen Datensätze übrig. Datum oder Betrag fehlt.")
            return

    except Exception as e:
        crash_report(e, context={"Phase": "Normalisierung", "Spalten": df.columns.tolist()})
        return

    # FILTER
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
        st.info("Analyse starten.")
        return

    kunden_mask = df[c_kun].isin(sel_kunden) if sel_kunden else True
    mask = (
        (df[c_dat].dt.date >= date_range[0]) &
        (df[c_dat].dt.date <= date_range[1]) &
        kunden_mask
    )

    f_df = df.loc[mask].copy()
    if f_df.empty:
        st.warning("Keine Datensätze im Filter.")
        return

    today = pd.Timestamp(datetime.now().date())
    df_offen = f_df[f_df[c_pay].isna()].copy()
    df_paid  = f_df[~f_df[c_pay].isna()].copy()

    tabs = st.tabs(["📊 Performance", "🔴 Forderungen", "💎 Strategie", "🔍 Forensik", "🏦 Bank"])

    # TAB 1
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

    # TAB 2
    with tabs[1]:
        st.subheader("Forderungs Management")
        if df_offen.empty:
            st.info("Keine offenen Posten.")
        else:
            df_offen["Verzug"] = np.where(
                df_offen["_Fällig_Datum"].isna(),
                np.nan,
                (today - df_offen["_Fällig_Datum"]).dt.days
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
                    df_offen.dropna(subset=["_Fällig_Datum"])
                    .groupby("_Fällig_Datum", as_index=False)[c_bet].sum()
                )
                if df_pred.empty:
                    st.info("Keine Prognose möglich.")
                else:
                    df_pred["Size"] = df_pred[c_bet].abs().clip(lower=0.1)
                    st.plotly_chart(px.scatter(df_pred, x="_Fällig_Datum", y=c_bet, size="Size", title="Cash Inflow Prognose"), width="stretch")

            show_cols = [c_dat, "Fällig_Text", c_kun, c_nr, c_bet, "Verzug"]
            show_cols = [c for c in show_cols if c in df_offen.columns]
            st.dataframe(make_arrow_safe(df_offen[show_cols].sort_values("Verzug", ascending=False)), width="stretch")

            st.download_button(
                "📥 Excel OP Liste",
                data=to_excel_bytes(df_offen),
                file_name="OP_Liste.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                width="stretch"
            )

    # TAB 3
    with tabs[2]:
        st.subheader("ABC Analyse")
        if c_kun not in f_df.columns:
            st.info("Keine Kunden Spalte verfügbar.")
        else:
            abc = f_df.groupby(c_kun, as_index=False)[c_bet].sum().sort_values(c_bet, ascending=False)
            st.plotly_chart(px.bar(abc.head(15), x=c_kun, y=c_bet, title="Top Kunden"), width="stretch")
            top3_share = (abc[c_bet].head(3).sum() / float(f_df[c_bet].sum()) * 100) if float(f_df[c_bet].sum()) > 0 else 0
            st.metric("Klumpenrisiko Top 3", f"{top3_share:.1f}%")

    # TAB 4
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
                    if nums.nunique() <= 1:
                        st.info("Nicht genug Nummern für Prüfung.")
                    else:
                        arr = np.array(sorted(nums.unique()))
                        full = np.arange(arr.min(), arr.max() + 1)
                        miss = np.setdiff1d(full, arr)
                        if len(miss) == 0:
                            st.success("Lückenlos")
                        else:
                            st.warning(f"Nummern fehlen: {len(miss)}")
                            st.write(miss[:20])
                except Exception as e:
                    st.info("Nummernkreis nicht prüfbar.")

    # TAB 5
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
                crash_report(e, context={"Phase": "Bank Import"})

# -----------------------------
# RUN WITH GLOBAL EXCEPTION REPORTING
# -----------------------------
try:
    main()
except Exception as e:
    crash_report(e, context={"Phase": "Global", "Python": sys.version})
