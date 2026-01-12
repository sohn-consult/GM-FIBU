import streamlit as st
import pandas as pd
import numpy as np
import re
from io import BytesIO
from datetime import date
import PyPDF2

# =========================================================
# PAGE CONFIG + STYLE
# =========================================================
st.set_page_config(page_title="Sohn Consult Liquidität", page_icon="👔", layout="wide")

st.markdown(
    """
    <style>
    .kpiGrid { display: grid; grid-template-columns: repeat(4, minmax(0, 1fr)); gap: 12px; }
    .kpiCard {
        background: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 12px;
        padding: 14px 16px;
        box-shadow: 0 2px 6px rgba(0,0,0,0.05);
        overflow: visible;
    }
    .kpiLabel { font-size: 12px; color: #475569; margin-bottom: 6px; }
    .kpiValue {
        font-size: 18px;
        font-weight: 800;
        color: #0F172A;
        line-height: 1.2;
        word-break: break-word;
        white-space: normal;
    }
    .kpiSub { font-size: 12px; color: #64748B; margin-top: 6px; }
    @media (max-width: 1100px) {
        .kpiGrid { grid-template-columns: repeat(2, minmax(0, 1fr)); }
        .kpiValue { font-size: 16px; }
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("👔 Sohn Consult | Liquidität, Offene Posten, Bank Abgleich")
st.caption("Stabiler Berater Modus. Excel Import mit manuellem Mapping. Bank Abgleich per PDF Umsätze Druckansicht.")

# =========================================================
# FORMAT HELPERS
# =========================================================
MONEY_KEYS = ["betrag", "summe", "umsatz", "saldo", "invoice_betrag"]
DATE_KEYS = ["datum", "faellig", "fällig", "buchung", "wertstellung", "gezahlt"]

def format_eur(x) -> str:
    if x is None or pd.isna(x):
        return "0,00 €"
    return f"{float(x):,.2f} €".replace(",", "X").replace(".", ",").replace("X", ".")

def format_date_de(x) -> str:
    if x is None or pd.isna(x):
        return ""
    try:
        return pd.to_datetime(x).strftime("%d.%m.%Y")
    except Exception:
        return ""

def df_for_display(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    for c in out.columns:
        cl = str(c).lower()

        if pd.api.types.is_datetime64_any_dtype(out[c]) or any(k in cl for k in DATE_KEYS):
            out[c] = out[c].apply(format_date_de).astype("string")
            continue

        if any(k in cl for k in MONEY_KEYS):
            tmp = pd.to_numeric(out[c], errors="coerce")
            out[c] = tmp.apply(format_eur).astype("string")
            continue

        if pd.api.types.is_numeric_dtype(out[c]):
            out[c] = pd.to_numeric(out[c], errors="coerce")
        else:
            out[c] = out[c].astype("string")

    return out

def render_kpis(items: list[dict]):
    cards = []
    for it in items:
        label = it.get("label", "")
        value = it.get("value", "")
        sub = it.get("sub", "")
        cards.append(
            f"""
            <div class="kpiCard">
              <div class="kpiLabel">{label}</div>
              <div class="kpiValue">{value}</div>
              <div class="kpiSub">{sub}</div>
            </div>
            """
        )
    st.markdown(f'<div class="kpiGrid">{"".join(cards)}</div>', unsafe_allow_html=True)

def to_excel_bytes(df: pd.DataFrame, sheet_name: str = "Export") -> bytes:
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    return bio.getvalue()

# =========================================================
# PARSING HELPERS
# =========================================================
def make_unique_columns(cols: list) -> list:
    seen = {}
    out = []
    for c in cols:
        base = str(c).strip()
        if base == "" or base.lower() in ["nan", "none"]:
            base = "Spalte"
        if base not in seen:
            seen[base] = 1
            out.append(base)
        else:
            seen[base] += 1
            out.append(f"{base}_{seen[base]}")
    return out

def to_number_series(s: pd.Series) -> pd.Series:
    if pd.api.types.is_numeric_dtype(s):
        return pd.to_numeric(s, errors="coerce")
    x = s.astype("string").str.strip()
    x = x.str.replace("\u00a0", "", regex=False)
    x = x.str.replace("€", "", regex=False)
    x = x.str.replace(" ", "", regex=False)
    x = x.str.replace(".", "", regex=False)
    x = x.str.replace(",", ".", regex=False)
    return pd.to_numeric(x, errors="coerce")

def to_date_series(s: pd.Series) -> pd.Series:
    if pd.api.types.is_datetime64_any_dtype(s):
        return pd.to_datetime(s, errors="coerce")

    x = s.astype("string").str.strip()
    x = x.replace({"": pd.NA, "nan": pd.NA, "NaT": pd.NA})

    dt1 = pd.to_datetime(x, format="%d.%m.%Y", errors="coerce")
    dt2 = pd.to_datetime(x, format="%d.%m.%Y %H:%M", errors="coerce")
    dt3 = pd.to_datetime(x, format="%d.%m.%Y %H:%M:%S", errors="coerce")
    dt4 = pd.to_datetime(x, format="%Y-%m-%d", errors="coerce")
    dt5 = pd.to_datetime(x, format="%Y-%m-%d %H:%M", errors="coerce")
    dt6 = pd.to_datetime(x, format="%Y-%m-%d %H:%M:%S", errors="coerce")

    out = dt1.fillna(dt2).fillna(dt3).fillna(dt4).fillna(dt5).fillna(dt6)
    out = out.fillna(pd.to_datetime(x, errors="coerce"))
    return out

def guess_col(cols: list[str], keys: list[str]) -> str | None:
    for c in cols:
        cl = str(c).lower()
        if any(k in cl for k in keys):
            return c
    return None

# =========================================================
# EXCEL LOADING (NO MISALIGNED BOOLEAN INDEXERS)
# =========================================================
def read_excel_raw(file, sheet_name: str) -> pd.DataFrame:
    return pd.read_excel(file, sheet_name=sheet_name, header=None, engine="openpyxl")

def normalize_with_header_row(raw: pd.DataFrame, header_row: int) -> pd.DataFrame:
    if raw.empty:
        return raw

    header_row = max(0, min(int(header_row), len(raw) - 1))
    header = raw.iloc[header_row].astype("string").fillna("").tolist()
    cols = make_unique_columns([str(c).strip() for c in header])

    df = raw.iloc[header_row + 1 :].copy()
    df.columns = cols
    df.dropna(how="all", inplace=True)

    # Drop Unnamed columns using list (no boolean series index issues)
    keep_cols = [c for c in df.columns if not str(c).strip().lower().startswith("unnamed")]
    df = df[keep_cols].copy()

    # Remove fully empty columns
    empty_cols = [c for c in df.columns if df[c].isna().all()]
    if empty_cols:
        df = df.drop(columns=empty_cols)

    return df

def auto_find_header_row(raw: pd.DataFrame) -> int:
    keywords = ["kunde", "debitor", "re", "datum", "fällig", "faellig", "betrag", "brutto", "netto", "gezahlt", "eingang"]
    best_i = 0
    best_score = -1
    scan = min(len(raw), 50)
    for i in range(scan):
        row = raw.iloc[i].astype("string").fillna("").str.lower()
        score = 0
        for cell in row.tolist():
            if any(k in str(cell) for k in keywords):
                score += 1
        if score > best_score:
            best_score = score
            best_i = i
    return best_i

def build_invoice_table(df_norm: pd.DataFrame, col_kunde: str, col_re: str, col_redat: str, col_faellig: str, col_betrag: str, col_paid: str | None) -> pd.DataFrame:
    out = pd.DataFrame()
    out["Kunde"] = df_norm[col_kunde].astype("string") if col_kunde else pd.Series([""] * len(df_norm), dtype="string")
    out["RE_Nr"] = df_norm[col_re].astype("string") if col_re else pd.Series([""] * len(df_norm), dtype="string")
    out["RE_Datum"] = to_date_series(df_norm[col_redat]) if col_redat else pd.NaT
    out["Faellig"] = to_date_series(df_norm[col_faellig]) if col_faellig else pd.NaT
    out["Betrag"] = to_number_series(df_norm[col_betrag]) if col_betrag else 0.0
    out["Gezahlt_Am"] = to_date_series(df_norm[col_paid]) if col_paid else pd.NaT

    out["Kunde"] = out["Kunde"].fillna("").astype("string").str.strip()
    out["RE_Nr"] = out["RE_Nr"].fillna("").astype("string").str.strip()
    out["Betrag"] = pd.to_numeric(out["Betrag"], errors="coerce").fillna(0.0)

    out = out.dropna(subset=["RE_Datum"]).copy()
    return out

# =========================================================
# BANK PDF PARSER
# =========================================================
TX_RE = re.compile(r"\)\s*(\d{2}\.\d{2}\.\d{4})\s*(\d{2}\.\d{2}\.\d{4})\s*([+-]?\d{1,3}(?:\.\d{3})*,\d{2})")
INVOICE_NO_RE = re.compile(r"\b(20\d{6,}[-/]\S+|\d{6,})\b")

def parse_bank_pdf(file) -> pd.DataFrame:
    reader = PyPDF2.PdfReader(file)
    lines = []
    for p in reader.pages:
        t = p.extract_text() or ""
        for ln in t.splitlines():
            ln = ln.replace("\u00a0", " ").strip()
            if ln:
                lines.append(ln)

    def is_meta(ln: str) -> bool:
        l = ln.lower()
        return any(k in l for k in ["umsätze - druckansicht", "umsätze vom", "kontostand", "buchungwertstellung", "sichteinlagen", "iban"])

    tx = []
    party = ""
    desc_buf = []

    for ln in lines:
        if is_meta(ln):
            continue

        m = TX_RE.search(ln)
        if m:
            buchung = pd.to_datetime(m.group(1), format="%d.%m.%Y", errors="coerce")
            wert = pd.to_datetime(m.group(2), format="%d.%m.%Y", errors="coerce")

            amt_raw = m.group(3).strip()
            sign = -1.0 if amt_raw.startswith("-") else 1.0
            amt = to_number_series(pd.Series([amt_raw.lstrip("+-")])).iloc[0]
            amt = float(amt) * sign if not pd.isna(amt) else np.nan

            tx.append(
                {
                    "Buchung": buchung,
                    "Wertstellung": wert,
                    "Betrag": amt,
                    "Gegenpartei": party.strip(),
                    "Verwendungszweck": " ".join(desc_buf).strip(),
                }
            )
            desc_buf = []
            continue

        if len(ln) <= 60 and not ln.startswith("(") and re.search(r"[A-Za-zÄÖÜäöüß]", ln):
            party = ln
            desc_buf = []
            continue

        desc_buf.append(ln)

    df = pd.DataFrame(tx)
    if df.empty:
        return df

    df["Buchung"] = pd.to_datetime(df["Buchung"], errors="coerce")
    df["Wertstellung"] = pd.to_datetime(df["Wertstellung"], errors="coerce")
    df["Betrag"] = pd.to_numeric(df["Betrag"], errors="coerce")
    return df

def reconcile_bank_vs_open(bank: pd.DataFrame, open_inv: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    if bank.empty or open_inv.empty:
        return pd.DataFrame(), bank

    bank_in = bank[bank["Betrag"] > 0].copy()
    inv = open_inv.copy()

    inv["RE_Nr_key"] = inv["RE_Nr"].astype("string").str.strip()
    inv["Kunde_key"] = inv["Kunde"].astype("string").str.lower().str.strip()

    inv_map = {}
    for idx, r in inv.iterrows():
        key = str(r["RE_Nr_key"])
        if key and key.lower() != "nan":
            inv_map.setdefault(key, []).append(idx)

    used_inv = set()
    matches = []

    # Match by invoice number
    for _, br in bank_in.iterrows():
        text = f"{br.get('Gegenpartei','')} {br.get('Verwendungszweck','')}".strip()
        tokens = [str(t).strip() for t in INVOICE_NO_RE.findall(text)]
        chosen = None
        for token in tokens:
            if token in inv_map:
                for iidx in inv_map[token]:
                    if iidx not in used_inv:
                        chosen = iidx
                        break
            if chosen is not None:
                break

        if chosen is not None:
            used_inv.add(chosen)
            ir = inv.loc[chosen]
            matches.append(
                {
                    "Buchung": br["Buchung"],
                    "Betrag": br["Betrag"],
                    "Gegenpartei": br.get("Gegenpartei", ""),
                    "Verwendungszweck": br.get("Verwendungszweck", ""),
                    "RE_Nr": ir["RE_Nr"],
                    "Kunde": ir["Kunde"],
                    "RE_Datum": ir["RE_Datum"],
                    "Faellig": ir["Faellig"],
                    "Invoice_Betrag": ir["Betrag"],
                    "MatchType": "RE Nummer",
                }
            )

    # Match by amount for remaining
    tolerance = 0.02
    still_open = inv.drop(index=list(used_inv), errors="ignore").copy()
    still_open["_amt"] = still_open["Betrag"].round(2)

    used_keys = set()
    for m in matches:
        k = f"{pd.to_datetime(m['Buchung'], errors='coerce').date()}|{round(float(m['Betrag']),2)}"
        used_keys.add(k)

    bank_rest = bank_in.copy()
    bank_rest["_key"] = bank_rest["Buchung"].dt.date.astype("string") + "|" + bank_rest["Betrag"].round(2).astype("string")
    bank_rest = bank_rest[~bank_rest["_key"].isin(used_keys)].copy()

    for _, br in bank_rest.iterrows():
        amt = round(float(br["Betrag"]), 2)
        cands = still_open[np.abs(still_open["_amt"] - amt) <= tolerance]
        if cands.empty:
            continue

        text = f"{br.get('Gegenpartei','')} {br.get('Verwendungszweck','')}".lower()
        cands = cands.copy()
        cands["name_hit"] = cands["Kunde_key"].apply(lambda k: 1 if k and k != "nan" and k in text else 0)
        cands = cands.sort_values(["name_hit", "RE_Datum"], ascending=[False, True])

        chosen = None
        for iidx in cands.index.tolist():
            if iidx not in used_inv:
                chosen = iidx
                break
        if chosen is None:
            continue

        used_inv.add(chosen)
        ir = inv.loc[chosen]
        matches.append(
            {
                "Buchung": br["Buchung"],
                "Betrag": br["Betrag"],
                "Gegenpartei": br.get("Gegenpartei", ""),
                "Verwendungszweck": br.get("Verwendungszweck", ""),
                "RE_Nr": ir["RE_Nr"],
                "Kunde": ir["Kunde"],
                "RE_Datum": ir["RE_Datum"],
                "Faellig": ir["Faellig"],
                "Invoice_Betrag": ir["Betrag"],
                "MatchType": "Betrag",
            }
        )

    matches_df = pd.DataFrame(matches)

    if matches_df.empty:
        unmatched = bank_in.copy()
    else:
        bank_in2 = bank_in.copy()
        bank_in2["_key"] = bank_in2["Buchung"].dt.date.astype("string") + "|" + bank_in2["Betrag"].round(2).astype("string")
        matches_df["_key"] = pd.to_datetime(matches_df["Buchung"], errors="coerce").dt.date.astype("string") + "|" + pd.to_numeric(matches_df["Betrag"], errors="coerce").round(2).astype("string")
        used = set(matches_df["_key"].dropna().astype(str).tolist())
        unmatched = bank_in2[~bank_in2["_key"].isin(used)].drop(columns=["_key"], errors="ignore")

    return matches_df.drop(columns=["_key"], errors="ignore"), unmatched

# =========================================================
# SIDEBAR INPUTS
# =========================================================
with st.sidebar:
    st.header("Import")
    fibu_file = st.file_uploader("1 Excel oder CSV", type=["xlsx", "xls", "csv"])
    bank_pdf = st.file_uploader("2 Bank PDF", type=["pdf"])
    st.divider()
    show_rows = st.number_input("Max Zeilen", min_value=50, max_value=2000, value=300, step=50)

if not fibu_file:
    st.info("Bitte Excel oder CSV hochladen.")
    st.stop()

# =========================================================
# LOAD AND NORMALIZE SOURCE TABLE
# =========================================================
file_name = getattr(fibu_file, "name", "").lower()

if file_name.endswith(".csv"):
    src = pd.read_csv(fibu_file, sep=None, engine="python", dtype="string")
    src.columns = make_unique_columns([str(c).strip() for c in src.columns])
    df_norm = src.copy()
    header_row_used = 0
    sheet_used = "CSV"
else:
    xls = pd.ExcelFile(fibu_file)
    sheet_used = st.sidebar.selectbox("Sheet", options=xls.sheet_names, index=0)
    raw = read_excel_raw(fibu_file, sheet_used)
    auto_header = auto_find_header_row(raw)
    header_row_used = st.sidebar.number_input("Header Zeile (0 basiert)", min_value=0, max_value=max(len(raw) - 1, 0), value=int(auto_header), step=1)
    df_norm = normalize_with_header_row(raw, header_row_used)

if df_norm.empty:
    st.error("Import Ergebnis ist leer. Prüfe Sheet und Header Zeile.")
    st.stop()

st.subheader("Import Vorschau")
st.write(f"Quelle: {sheet_used} | Header Zeile: {header_row_used} | Shape: {df_norm.shape}")
with st.expander("Vorschau Tabelle", expanded=False):
    st.dataframe(df_for_display(df_norm.head(25)), width="stretch")

# =========================================================
# MANUAL COLUMN MAPPING (USER REQUEST)
# =========================================================
cols = list(df_norm.columns)

auto_kunde = guess_col(cols, ["kunde", "debitor", "name"])
auto_re = guess_col(cols, ["re_n", "re-n", "re nr", "rechnungs", "beleg", "nummer"])
auto_redat = guess_col(cols, ["re_datum", "re-datum", "belegdat", "datum"])
auto_faellig = guess_col(cols, ["fällig", "faellig", "termin", "fälligkeit", "faelligkeit"])
auto_betrag = guess_col(cols, ["betrag (netto)", "betrag netto", "betrag (brutto)", "betrag brutto", "betrag", "brutto", "netto", "summe", "umsatz"])
auto_paid = guess_col(cols, ["gezahlt am", "gezahlt", "eingang", "ausgleich", "zahlung"])

with st.sidebar:
    st.header("Mapping")
    col_kunde = st.selectbox("Kunde", options=["<leer>"] + cols, index=(["<leer>"] + cols).index(auto_kunde) if auto_kunde in cols else 0)
    col_re = st.selectbox("RE Nummer", options=["<leer>"] + cols, index=(["<leer>"] + cols).index(auto_re) if auto_re in cols else 0)
    col_redat = st.selectbox("Rechnungsdatum", options=cols, index=cols.index(auto_redat) if auto_redat in cols else 0)
    col_faellig = st.selectbox("Fälligkeit", options=["<leer>"] + cols, index=(["<leer>"] + cols).index(auto_faellig) if auto_faellig in cols else 0)
    col_betrag = st.selectbox("Betrag", options=cols, index=cols.index(auto_betrag) if auto_betrag in cols else 0)
    col_paid = st.selectbox("Zahldatum", options=["<keins>"] + cols, index=(["<keins>"] + cols).index(auto_paid) if auto_paid in cols else 0)

# Normalize mapping selections
col_kunde = None if col_kunde == "<leer>" else col_kunde
col_re = None if col_re == "<leer>" else col_re
col_faellig = None if col_faellig == "<leer>" else col_faellig
col_paid = None if col_paid == "<keins>" else col_paid

# Build invoice table
try:
    inv = build_invoice_table(df_norm, col_kunde, col_re, col_redat, col_faellig, col_betrag, col_paid)
except Exception as e:
    st.error(f"Mapping oder Parsing fehlgeschlagen: {e}")
    st.stop()

if inv.empty:
    st.error("Nach Mapping und Bereinigung keine Datensätze übrig. Prüfe Datums und Betrag Spalten.")
    st.stop()

# =========================================================
# FILTERS
# =========================================================
inv["Monat"] = inv["RE_Datum"].dt.to_period("M").astype("string")
customers = sorted([c for c in inv["Kunde"].dropna().astype(str).unique().tolist() if c.strip()])

c1, c2 = st.columns([2, 2])
with c1:
    sel_customers = st.multiselect("Kunden Filter", options=customers, default=customers if customers else [])
with c2:
    min_d = inv["RE_Datum"].min().date()
    max_d = inv["RE_Datum"].max().date()
    dr = st.date_input("Zeitraum", value=(min_d, max_d))

d_from, d_to = (dr[0], dr[1]) if isinstance(dr, tuple) and len(dr) == 2 else (min_d, max_d)

f = inv.copy()
f = f[(f["RE_Datum"].dt.date >= d_from) & (f["RE_Datum"].dt.date <= d_to)]
if sel_customers:
    f = f[f["Kunde"].isin(sel_customers)]

today = pd.Timestamp(date.today())
f["Offen"] = f["Gezahlt_Am"].isna()
offen = f[f["Offen"]].copy()
bezahlt = f[~f["Offen"]].copy()

offen["VerzugTage"] = np.where(offen["Faellig"].notna(), (today - offen["Faellig"]).dt.days, np.nan)

def aging_bucket(v):
    if pd.isna(v):
        return "Unbekannt"
    if v <= 0:
        return "Pünktlich"
    if v <= 30:
        return "1 bis 30"
    if v <= 60:
        return "31 bis 60"
    return "größer 60"

offen["Aging"] = offen["VerzugTage"].apply(aging_bucket)

# KPIs
rev = float(f["Betrag"].sum())
op_sum = float(offen["Betrag"].sum()) if not offen.empty else 0.0
overdue_sum = float(offen.loc[offen["VerzugTage"] > 0, "Betrag"].sum()) if not offen.empty else 0.0

dso = None
if not bezahlt.empty:
    dso_val = (bezahlt["Gezahlt_Am"] - bezahlt["RE_Datum"]).dt.days.mean()
    dso = float(dso_val) if pd.notna(dso_val) else None

# =========================================================
# EXECUTIVE OVERVIEW
# =========================================================
st.markdown("## Executive Überblick")
render_kpis(
    [
        {"label": "Umsatz im Zeitraum", "value": format_eur(rev), "sub": f"Zeitraum {format_date_de(d_from)} bis {format_date_de(d_to)}"},
        {"label": "Offene Posten", "value": format_eur(op_sum), "sub": f"Offene Belege {int(len(offen))}"},
        {"label": "Überfällig", "value": format_eur(overdue_sum), "sub": "Fälligkeit überschritten"},
        {"label": "Zahlungsdauer Durchschnitt", "value": (f"{dso:.1f} Tage" if dso and dso > 0 else "Nicht verfügbar"), "sub": "Nur bei vorhandenem Zahldatum"},
    ]
)

st.divider()

# =========================================================
# TABS
# =========================================================
tab1, tab2, tab3 = st.tabs(["Offene Posten", "Kunden Fokus", "Bank Abgleich"])

with tab1:
    st.subheader("Offene Posten und Aging")
    aging_sum = offen.groupby("Aging")["Betrag"].sum().reset_index()
    st.dataframe(df_for_display(aging_sum.rename(columns={"Betrag": "Betrag"})), width="stretch")

    show_cols = ["Kunde", "RE_Nr", "RE_Datum", "Faellig", "Betrag", "VerzugTage", "Aging"]
    show = offen.sort_values(["VerzugTage", "Faellig"], ascending=[False, True]).copy()
    st.dataframe(df_for_display(show[show_cols].head(int(show_rows))), width="stretch")

    st.download_button(
        "Excel Export OP Liste",
        data=to_excel_bytes(offen[show_cols], sheet_name="OP"),
        file_name="OP_Liste.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

with tab2:
    st.subheader("Kunden Fokus")
    by_cust = f.groupby("Kunde")["Betrag"].sum().sort_values(ascending=False).reset_index()
    st.dataframe(df_for_display(by_cust.head(25)), width="stretch")

    top3_share = (float(by_cust["Betrag"].head(3).sum()) / rev * 100) if rev > 0 else 0.0
    render_kpis(
        [
            {"label": "Klumpenrisiko Top 3", "value": f"{top3_share:.1f} %", "sub": "Anteil am Umsatz im Zeitraum"},
            {"label": "Kunden im Filter", "value": f"{len(by_cust)}", "sub": "Anzahl Kunden mit Betrag"},
            {"label": "Offene Posten", "value": format_eur(op_sum), "sub": "Summe offen"},
            {"label": "Überfällig", "value": format_eur(overdue_sum), "sub": "Summe überfällig"},
        ]
    )

    st.download_button(
        "Excel Export Kundenumsatz",
        data=to_excel_bytes(by_cust, sheet_name="Kunden"),
        file_name="Kunden_Umsatz.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

with tab3:
    st.subheader("Bank Abgleich")

    if not bank_pdf:
        st.info("Bitte Bank PDF hochladen. Erwartet wird das Format Umsätze Druckansicht.")
    else:
        bank = parse_bank_pdf(bank_pdf)

        if bank.empty:
            st.warning("Keine Buchungen erkannt. Wenn das PDF ein Scan ist, braucht es OCR. Diese Version nutzt kein OCR.")
        else:
            st.write("Erkannte Buchungen Auszug")
            st.dataframe(df_for_display(bank[["Buchung", "Wertstellung", "Betrag", "Gegenpartei", "Verwendungszweck"]].head(80)), width="stretch")

            matches, unmatched = reconcile_bank_vs_open(bank, offen)

            total_in = float(bank.loc[bank["Betrag"] > 0, "Betrag"].sum())
            render_kpis(
                [
                    {"label": "Bank Eingänge gesamt", "value": format_eur(total_in), "sub": "Positiver Betrag"},
                    {"label": "Zuordnung Treffer", "value": f"{int(len(matches))}", "sub": "RE Nummer oder Betrag"},
                    {"label": "Nicht zugeordnet", "value": f"{int(len(unmatched))}", "sub": "Manuelle Prüfung"},
                    {"label": "Offene Posten Summe", "value": format_eur(op_sum), "sub": "Nach Import"},
                ]
            )

            st.divider()
            st.write("Zuordnungen Bank zu Offene Posten")
            if matches.empty:
                st.info("Keine Matches gefunden. Typisch: keine RE Nummer im Text oder Betrag weicht ab.")
            else:
                mcols = ["Buchung", "Betrag", "Gegenpartei", "RE_Nr", "Kunde", "Invoice_Betrag", "MatchType"]
                st.dataframe(df_for_display(matches[mcols].head(int(show_rows))), width="stretch")

            st.write("Nicht zugeordnete Bank Eingänge")
            if unmatched.empty:
                st.success("Keine offenen Bank Eingänge ohne Zuordnung.")
            else:
                ucols = ["Buchung", "Betrag", "Gegenpartei", "Verwendungszweck"]
                st.dataframe(df_for_display(unmatched[ucols].head(int(show_rows))), width="stretch")

            cexp1, cexp2 = st.columns(2)
            with cexp1:
                st.download_button(
                    "Excel Export Matches",
                    data=to_excel_bytes(matches, sheet_name="Matches"),
                    file_name="Bank_Matches.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            with cexp2:
                st.download_button(
                    "Excel Export Nicht zugeordnet",
                    data=to_excel_bytes(unmatched, sheet_name="Unmatched"),
                    file_name="Bank_Unmatched.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
