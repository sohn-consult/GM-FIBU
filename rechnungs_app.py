import streamlit as st
import pandas as pd
import numpy as np
import re
from io import BytesIO
from datetime import datetime, timedelta

st.set_page_config(page_title="Sohn Consult Liquidität Scanner", page_icon="👔", layout="wide")

st.title("Sohn Consult Liquidität Scanner")
st.caption("Stable Core. Berater Modus für schnellen Überblick über Debitoren, offene Posten, Liquiditätsindikatoren.")

# -----------------------------
# Helpers
# -----------------------------
KEYS = ["kunde", "debitor", "re", "nr", "datum", "fällig", "faellig", "betrag", "brutto", "netto", "gezahlt", "zahlung", "eingang", "ausgleich"]

def read_excel_raw(file) -> pd.DataFrame:
    return pd.read_excel(file, header=None, engine="openpyxl")

def score_row_as_header(row: pd.Series) -> int:
    vals = row.astype(str).str.lower().str.strip()
    hits = 0
    for v in vals:
        if v in ("nan", "none", ""):
            continue
        if any(k in v for k in KEYS):
            hits += 1
    return hits

def pick_header_row(df_raw: pd.DataFrame, max_rows: int = 25) -> int:
    best_i = 0
    best_score = -1
    upto = min(max_rows, len(df_raw))
    for i in range(upto):
        s = score_row_as_header(df_raw.iloc[i])
        if s > best_score:
            best_score = s
            best_i = i
    return best_i

def normalize_table(df_raw: pd.DataFrame) -> pd.DataFrame:
    h = pick_header_row(df_raw)
    header = df_raw.iloc[h].astype(str).str.strip()

    cols = []
    for i, v in enumerate(header.tolist()):
        vl = v.lower()
        if v and vl not in ("nan", "none") and any(k in vl for k in KEYS):
            cols.append(v)
        else:
            cols.append(f"col_{i:02d}")

    df = df_raw.iloc[h+1:].copy()
    df.columns = cols
    df.dropna(how="all", inplace=True)

    # remove columns fully empty
    empty_cols = [c for c in df.columns if df[c].isna().all()]
    if empty_cols:
        df = df.drop(columns=empty_cols)

    # dedupe column names
    df = df.loc[:, ~df.columns.duplicated()]

    return df

def to_num(series: pd.Series) -> pd.Series:
    x = series.astype(str).str.strip()
    x = x.str.replace("€", "", regex=False).str.replace(" ", "", regex=False)
    x = x.str.replace(".", "", regex=False).str.replace(",", ".", regex=False)
    return pd.to_numeric(x, errors="coerce")

def to_date(series: pd.Series) -> pd.Series:
    x = series.astype(str).str.strip()
    out = pd.Series(pd.NaT, index=series.index, dtype="datetime64[ns]")

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

def euro(x: float) -> str:
    if x is None or pd.isna(x):
        return "0,00 €"
    return f"{float(x):,.2f} €".replace(",", "X").replace(".", ",").replace("X", ".")

def find_col(cols, patterns):
    for c in cols:
        cl = str(c).lower()
        if any(p in cl for p in patterns):
            return c
    return None

def to_excel_bytes(df: pd.DataFrame, sheet_name: str = "Analyse") -> bytes:
    out = BytesIO()
    with pd.ExcelWriter(out) as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    return out.getvalue()

def aging_bucket(days_overdue):
    if pd.isna(days_overdue):
        return "Unbekannt"
    if days_overdue <= 0:
        return "0 Pünktlich"
    if days_overdue <= 30:
        return "1 1-30"
    if days_overdue <= 60:
        return "2 31-60"
    return "3 >60"

def week_start(dt: pd.Timestamp) -> pd.Timestamp:
    return dt - pd.Timedelta(days=dt.weekday())

# -----------------------------
# Upload
# -----------------------------
f = st.file_uploader("Excel Datei laden (XLSX)", type=["xlsx"])
if not f:
    st.stop()

raw = read_excel_raw(f)
df = normalize_table(raw)

st.subheader("Import Ergebnis")
st.write("Datei:", f.name)
st.write("Shape:", df.shape)
st.write("Spalten (erste 30):", list(df.columns)[:30])

with st.expander("Vorschau Rohdaten", expanded=False):
    st.write(df.head(25))

# -----------------------------
# Mapping (Auto + Override)
# -----------------------------
cols = list(df.columns)

auto_kunde = find_col(cols, ["kunde", "debitor", "name"])
auto_nr    = find_col(cols, ["re-nr", "re nr", "nummer", "beleg"])
auto_rdat  = find_col(cols, ["re-datum", "re datum", "datum"])
auto_fae   = find_col(cols, ["fällig", "faellig", "termin"])
auto_bet   = find_col(cols, ["betrag", "brutto", "netto", "summe"])
auto_pay   = find_col(cols, ["gezahlt", "eingang", "zahlung", "ausgleich"])

st.sidebar.header("Mapping")
c_kunde = st.sidebar.selectbox("Kunde", cols, index=cols.index(auto_kunde) if auto_kunde in cols else 0)
c_nr    = st.sidebar.selectbox("Re Nummer", cols, index=cols.index(auto_nr) if auto_nr in cols else 0)
c_rdat  = st.sidebar.selectbox("Rechnungsdatum", cols, index=cols.index(auto_rdat) if auto_rdat in cols else 0)
c_fae   = st.sidebar.selectbox("Fälligkeit", cols, index=cols.index(auto_fae) if auto_fae in cols else 0)
c_bet   = st.sidebar.selectbox("Betrag", cols, index=cols.index(auto_bet) if auto_bet in cols else 0)
c_pay   = st.sidebar.selectbox("Zahldatum (optional)", ["<keins>"] + cols, index=0 if auto_pay is None else (["<keins>"] + cols).index(auto_pay))

# -----------------------------
# Normalisierung
# -----------------------------
work = df.copy()

work[c_kunde] = work[c_kunde].astype(str).str.strip()
work[c_nr] = work[c_nr].astype(str).str.strip()
work[c_bet] = to_num(work[c_bet])
work[c_rdat] = to_date(work[c_rdat])
work["_Faellig"] = to_date(work[c_fae])

if c_pay != "<keins>":
    work["_Pay"] = to_date(work[c_pay])
else:
    work["_Pay"] = pd.NaT

# Filter: brauchbare Datensätze
work = work.dropna(subset=[c_rdat, c_bet]).copy()

today = pd.Timestamp(datetime.now().date())

# offen = kein Pay Datum
offen = work[work["_Pay"].isna()].copy()
bezahlt = work[~work["_Pay"].isna()].copy()

# Verzug nur wenn Fälligkeit parsebar
offen["_OverdueDays"] = np.where(offen["_Faellig"].isna(), np.nan, (today - offen["_Faellig"]).dt.days)
offen["_Bucket"] = offen["_OverdueDays"].apply(aging_bucket)

# -----------------------------
# Berater Cockpit
# -----------------------------
st.markdown("## Executive Überblick")

k1, k2, k3, k4, k5 = st.columns(5)

total_sum = float(work[c_bet].sum())
open_sum = float(offen[c_bet].sum()) if len(offen) else 0.0
overdue_sum = float(offen.loc[offen["_OverdueDays"] > 0, c_bet].sum()) if len(offen) else 0.0
count_docs = int(len(work))
count_open = int(len(offen))

k1.metric("Gesamt Debitoren", euro(total_sum))
k2.metric("Offene Posten", euro(open_sum))
k3.metric("Überfällig", euro(overdue_sum))
k4.metric("Belege", f"{count_docs}")
k5.metric("Offen", f"{count_open}")

if len(bezahlt) > 0:
    dso = (bezahlt["_Pay"] - bezahlt[c_rdat]).dt.days.mean()
    st.write("Ø Zahlungsdauer Tage:", round(float(dso), 1))
else:
    st.write("Ø Zahlungsdauer Tage: nicht verfügbar (kein Zahldatum)")

# -----------------------------
# Tabs
# -----------------------------
tabs = st.tabs(["Offene Posten", "Aging", "Top Schuldner", "Cash Forecast", "Export"])

# Offene Posten
with tabs[0]:
    st.subheader("Offene Posten Liste")
    show_cols = [c_kunde, c_nr, c_rdat, c_fae, c_bet]
    view = offen.copy()
    view["Verzug Tage"] = view["_OverdueDays"]
    view = view.sort_values("Verzug Tage", ascending=False)
    st.write(view[show_cols + ["Verzug Tage"]].head(50))

# Aging
with tabs[1]:
    st.subheader("Aging Übersicht")
    if len(offen) == 0:
        st.info("Keine offenen Posten.")
    else:
        aging = offen.groupby("_Bucket", dropna=False)[c_bet].sum().reset_index()
        aging = aging.sort_values("_Bucket")
        st.write(aging)
        st.write("Hinweis: Unbekannt bedeutet Fälligkeit nicht als Datum erkennbar.")

# Top Schuldner
with tabs[2]:
    st.subheader("Top Schuldner nach offenem Betrag")
    if len(offen) == 0:
        st.info("Keine offenen Posten.")
    else:
        top = offen.groupby(c_kunde)[c_bet].sum().reset_index().sort_values(c_bet, ascending=False).head(15)
        st.write(top)

# Cash Forecast
with tabs[3]:
    st.subheader("Cash In Forecast nach Fälligkeit")
    pred = offen.dropna(subset=["_Faellig"]).copy()
    if len(pred) == 0:
        st.info("Keine Fälligkeiten als Datum erkennbar.")
    else:
        mode = st.radio("Aggregation", ["Woche", "Monat"], horizontal=True)
        if mode == "Woche":
            pred["Periode"] = pred["_Faellig"].apply(week_start)
            fc = pred.groupby("Periode")[c_bet].sum().reset_index().sort_values("Periode")
        else:
            pred["Periode"] = pred["_Faellig"].dt.to_period("M").astype(str)
            fc = pred.groupby("Periode")[c_bet].sum().reset_index()
        st.write(fc)

# Export
with tabs[4]:
    st.subheader("Export")
    export_offen = offen.copy()
    export_offen["Verzug Tage"] = export_offen["_OverdueDays"]
    export_offen["Aging"] = export_offen["_Bucket"]

    st.download_button(
        "OP Liste als Excel",
        data=to_excel_bytes(export_offen[[c_kunde, c_nr, c_rdat, c_fae, c_bet, "Verzug Tage", "Aging"]].copy(), "OP"),
        file_name="OP_Liste.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

    st.download_button(
        "Gesamtdaten als Excel",
        data=to_excel_bytes(work.copy(), "Daten"),
        file_name="Daten_normalisiert.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

st.success("Stabiler Berater Workflow aktiv. Wenn du möchtest, erweitern wir als nächsten Schritt den Bank CSV Abgleich.")
