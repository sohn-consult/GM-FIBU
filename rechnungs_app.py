import streamlit as st
import pandas as pd
import numpy as np
import re
from io import BytesIO

st.set_page_config(page_title="Sohn Consult BI Safe Mode", page_icon="👔", layout="wide")

st.title("Sohn Consult BI Safe Mode")
st.caption("Minimal, robust, ohne Arrow Tabellen, ohne Plotly. Ziel ist Stabilität.")

# -----------------------------
# Helpers
# -----------------------------
KEYS = ["kunde", "debitor", "re", "nr", "datum", "fällig", "faellig", "betrag", "brutto", "netto", "gezahlt", "zahlung"]

def read_excel_raw(file) -> pd.DataFrame:
    # Always read without header to avoid Excel multi header chaos
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
    # Find the best header row in first N rows
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
    # choose header row
    h = pick_header_row(df_raw)
    header = df_raw.iloc[h].astype(str).str.strip()

    # build column names: keep header like cells, else col_XX
    cols = []
    for i, v in enumerate(header.tolist()):
        vl = v.lower()
        if v and vl not in ("nan", "none") and any(k in vl for k in KEYS):
            cols.append(v)
        else:
            cols.append(f"col_{i:02d}")

    df = df_raw.iloc[h+1:].copy()
    df.columns = cols

    # drop fully empty rows and empty columns
    df.dropna(how="all", inplace=True)
    df = df.loc[:, ~df.columns.duplicated()]

    # remove columns that are completely empty
    empty_cols = [c for c in df.columns if df[c].isna().all()]
    if empty_cols:
        df = df.drop(columns=empty_cols)

    return df

def to_num(series: pd.Series) -> pd.Series:
    x = series.astype(str).str.strip()
    x = x.str.replace("€", "", regex=False).str.replace(" ", "", regex=False)
    x = x.str.replace(".", "", regex=False).str.replace(",", ".", regex=False)
    return pd.to_numeric(x, errors="coerce")

def to_date(series: pd.Series) -> pd.Series:
    # very conservative: accept only ISO or DE dot; else NaT
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

def find_col(cols, patterns):
    for c in cols:
        cl = str(c).lower()
        if any(p in cl for p in patterns):
            return c
    return None

# -----------------------------
# UI
# -----------------------------
f = st.file_uploader("Excel Datei laden (XLSX)", type=["xlsx"])
if not f:
    st.stop()

try:
    raw = read_excel_raw(f)
except Exception as e:
    st.error("Excel konnte nicht gelesen werden.")
    st.write(str(e))
    st.stop()

st.subheader("Import Status")
st.write("Rohdaten Shape:", raw.shape)

try:
    df = normalize_table(raw)
except Exception as e:
    st.error("Konnte Tabelle nicht normalisieren.")
    st.write(str(e))
    st.write("Rohdaten Vorschau:")
    st.write(raw.head(30))
    st.stop()

st.write("Normalisierte Tabelle Shape:", df.shape)
st.write("Spalten (erste 40):", list(df.columns)[:40])

# Show preview WITHOUT st.dataframe (Arrow avoidance)
st.subheader("Vorschau (ohne Arrow)")
st.write(df.head(20))

# Mapping auto
cols = list(df.columns)
c_kun = find_col(cols, ["kunde", "debitor", "name"])
c_nr  = find_col(cols, ["re-nr", "re nr", "nummer", "beleg"])
c_dat = find_col(cols, ["re-datum", "re datum", "datum"])
c_fae = find_col(cols, ["fällig", "faellig", "termin"])
c_bet = find_col(cols, ["betrag", "brutto", "netto", "summe"])
c_pay = find_col(cols, ["gezahlt", "eingang", "zahlung", "ausgleich"])

st.subheader("Auto Mapping Ergebnis")
st.write({
    "Kunde": c_kun,
    "RE Nummer": c_nr,
    "Rechnungsdatum": c_dat,
    "Fälligkeit": c_fae,
    "Betrag": c_bet,
    "Zahldatum": c_pay
})

# Minimal metrics only if mandatory cols exist
if c_dat and c_bet:
    d = df.copy()

    d[c_bet] = to_num(d[c_bet])
    d[c_dat] = to_date(d[c_dat])
    if c_pay:
        d[c_pay] = to_date(d[c_pay])

    d = d.dropna(subset=[c_dat, c_bet])
    st.subheader("Kennzahlen")
    st.write("Datensätze:", int(len(d)))
    st.write("Gesamtbetrag:", float(d[c_bet].sum()))

    if c_pay:
        paid = d.dropna(subset=[c_pay]).copy()
        if len(paid) > 0:
            dso = (paid[c_pay] - paid[c_dat]).dt.days.mean()
            st.write("Ø Zahlungsdauer Tage:", float(dso))
else:
    st.warning("Für Kennzahlen fehlen Rechnungsdatum oder Betrag Spalte. Mapping prüfen.")

st.success("Safe Mode erfolgreich. Wenn das läuft, können wir Features schrittweise wieder aktivieren.")
