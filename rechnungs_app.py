import streamlit as st
import pandas as pd
import numpy as np
import re
from io import BytesIO
from datetime import date
import PyPDF2
import plotly.express as px

# =========================================================
# PAGE CONFIG + STYLE
# =========================================================
st.set_page_config(page_title="Sohn Consult | Executive Cash BI", page_icon="👔", layout="wide")

st.markdown(
    """
    <style>
    :root{
      --bg:#F8FAFC;
      --card:#FFFFFF;
      --line:#E2E8F0;
      --text:#0F172A;
      --muted:#64748B;
      --accent:#1E3A8A;
      --ok:#16A34A;
      --warn:#F59E0B;
      --bad:#DC2626;
    }
    .stApp { background: var(--bg); }
    section[data-testid="stSidebar"] { background: #F1F5F9; border-right: 1px solid var(--line); }

    .topbar {
      background: linear-gradient(90deg, rgba(30,58,138,0.10), rgba(59,130,246,0.08));
      border: 1px solid var(--line);
      border-radius: 14px;
      padding: 14px 16px;
      display:flex;
      justify-content:space-between;
      gap:16px;
      align-items:flex-start;
      margin-bottom: 10px;
    }
    .brand { font-weight: 900; color: var(--accent); font-size: 16px; letter-spacing:0.2px; }
    .sub { color: var(--muted); font-size: 12px; margin-top:4px; }
    .pillRow { display:flex; gap:8px; flex-wrap:wrap; justify-content:flex-end; }
    .pill {
      border:1px solid var(--line);
      background: rgba(255,255,255,0.7);
      padding:6px 10px;
      border-radius: 999px;
      font-size: 12px;
      color: var(--text);
      white-space:nowrap;
    }

    .kpiGrid { display:grid; grid-template-columns: repeat(4, minmax(0, 1fr)); gap: 12px; }
    .kpiCard{
      background: var(--card);
      border: 1px solid var(--line);
      border-radius: 14px;
      padding: 14px 14px;
      box-shadow: 0 2px 8px rgba(0,0,0,0.05);
      position:relative;
      overflow: hidden;
    }
    .kpiStripe{
      position:absolute; top:0; left:0; right:0; height:4px;
      background: var(--accent);
    }
    .kpiLabel { font-size:12px; color: var(--muted); margin-top:6px; }
    .kpiValue { font-size: 20px; font-weight: 900; color: var(--text); line-height:1.15; }
    .kpiDelta { font-size: 12px; margin-top:6px; color: var(--muted); }

    .badge { display:inline-block; padding:4px 8px; border-radius:999px; font-size:12px; border:1px solid var(--line); }
    .bOK { color: var(--ok); background: rgba(22,163,74,0.08); border-color: rgba(22,163,74,0.25); }
    .bWARN { color: var(--warn); background: rgba(245,158,11,0.10); border-color: rgba(245,158,11,0.25); }
    .bBAD { color: var(--bad); background: rgba(220,38,38,0.08); border-color: rgba(220,38,38,0.25); }

    .insights{
      background: var(--card);
      border: 1px solid var(--line);
      border-radius: 14px;
      padding: 12px 14px;
      box-shadow: 0 2px 8px rgba(0,0,0,0.05);
    }
    .insTitle { font-weight: 900; color: var(--text); margin-bottom: 6px; }
    .insLine { color: var(--text); font-size: 13px; margin: 6px 0; }
    .insMuted { color: var(--muted); font-size: 12px; margin-top: 8px; }

    @media (max-width: 1100px) {
      .kpiGrid { grid-template-columns: repeat(2, minmax(0, 1fr)); }
      .kpiValue { font-size: 18px; }
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# =========================================================
# HELPERS
# =========================================================
def format_date_de(x) -> str:
    if x is None or pd.isna(x):
        return ""
    try:
        return pd.to_datetime(x).strftime("%d.%m.%Y")
    except Exception:
        return ""

def eur_full(x: float) -> str:
    if x is None or pd.isna(x):
        return "0,00 €"
    return f"{float(x):,.2f} €".replace(",", "X").replace(".", ",").replace("X", ".")

def eur_short(x: float) -> str:
    if x is None or pd.isna(x):
        return "0 €"
    v = float(x)
    sign = "-" if v < 0 else ""
    v = abs(v)
    if v >= 1_000_000:
        return f"{sign}{v/1_000_000:.2f} Mio €".replace(".", ",")
    if v >= 1_000:
        return f"{sign}{v/1_000:.0f} Tsd €"
    return f"{sign}{v:.0f} €"

def to_excel_bytes(df: pd.DataFrame, sheet_name: str = "Export") -> bytes:
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    return bio.getvalue()

def df_for_display(df: pd.DataFrame, money_cols=None, date_cols=None) -> pd.DataFrame:
    out = df.copy()
    money_cols = money_cols or []
    date_cols = date_cols or []
    for c in out.columns:
        if c in date_cols:
            out[c] = out[c].apply(format_date_de).astype("string")
        elif c in money_cols:
            tmp = pd.to_numeric(out[c], errors="coerce")
            out[c] = tmp.apply(eur_full).astype("string")
        else:
            if pd.api.types.is_datetime64_any_dtype(out[c]):
                out[c] = out[c].apply(format_date_de).astype("string")
            elif pd.api.types.is_numeric_dtype(out[c]):
                out[c] = pd.to_numeric(out[c], errors="coerce")
            else:
                out[c] = out[c].astype("string")
    return out

def badge_html(kind: str, text: str) -> str:
    klass = "bOK" if kind == "ok" else "bWARN" if kind == "warn" else "bBAD"
    return f'<span class="badge {klass}">{text}</span>'

def render_kpis(items: list[dict]):
    cards = []
    for it in items:
        cards.append(
            f"""
            <div class="kpiCard">
              <div class="kpiStripe"></div>
              <div class="kpiLabel">{it.get("label","")}</div>
              <div class="kpiValue">{it.get("value","")}</div>
              <div class="kpiDelta">{it.get("delta","")}</div>
            </div>
            """
        )
    st.markdown(f'<div class="kpiGrid">{"".join(cards)}</div>', unsafe_allow_html=True)

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
# EXCEL LOADING (STABLE)
# =========================================================
def read_excel_raw(file, sheet_name: str) -> pd.DataFrame:
    return pd.read_excel(file, sheet_name=sheet_name, header=None, engine="openpyxl")

def auto_find_header_row(raw: pd.DataFrame) -> int:
    keywords = ["kunde", "debitor", "name", "re", "datum", "fällig", "faellig", "betrag", "brutto", "netto", "gezahlt", "eingang"]
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

def normalize_with_header_row(raw: pd.DataFrame, header_row: int) -> pd.DataFrame:
    if raw.empty:
        return raw
    header_row = max(0, min(int(header_row), len(raw) - 1))
    header = raw.iloc[header_row].astype("string").fillna("").tolist()
    cols = make_unique_columns([str(c).strip() for c in header])

    df = raw.iloc[header_row + 1 :].copy()
    df.columns = cols
    df.dropna(how="all", inplace=True)

    keep_cols = [c for c in df.columns if not str(c).strip().lower().startswith("unnamed")]
    df = df[keep_cols].copy()

    empty_cols = [c for c in df.columns if df[c].isna().all()]
    if empty_cols:
        df = df.drop(columns=empty_cols)

    return df.reset_index(drop=True)

def build_invoice_table(df_norm: pd.DataFrame, col_kunde, col_re, col_redat, col_faellig, col_betrag, col_paid) -> pd.DataFrame:
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

    out = out.dropna(subset=["RE_Datum"]).copy().reset_index(drop=True)
    return out

# =========================================================
# BANK PDF PARSER + RECONCILIATION
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
    return df.reset_index(drop=True)

def reconcile_bank_vs_open(bank: pd.DataFrame, open_inv: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    if bank.empty or open_inv.empty:
        return pd.DataFrame(), bank

    bank_in = bank[bank["Betrag"] > 0].copy().reset_index(drop=True)
    inv = open_inv.copy().reset_index(drop=True)

    inv["RE_Nr_key"] = inv["RE_Nr"].astype("string").str.strip()
    inv["Kunde_key"] = inv["Kunde"].astype("string").str.lower().str.strip()

    inv_map = {}
    for idx, r in inv.iterrows():
        key = str(r["RE_Nr_key"])
        if key and key.lower() != "nan":
            inv_map.setdefault(key, []).append(idx)

    used_inv = set()
    used_bank = set()
    matches = []

    for bidx, br in bank_in.iterrows():
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
            used_bank.add(bidx)
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

    tolerance = 0.02
    still_open = inv.drop(index=list(used_inv), errors="ignore").copy()
    if not still_open.empty:
        still_open["_amt"] = still_open["Betrag"].round(2)

    bank_rest = bank_in.drop(index=list(used_bank), errors="ignore").copy()
    for bidx, br in bank_rest.iterrows():
        if still_open.empty:
            break
        amt = round(float(br["Betrag"]), 2)
        cands = still_open[np.abs(still_open["_amt"] - amt) <= tolerance].copy()
        if cands.empty:
            continue

        text = f"{br.get('Gegenpartei','')} {br.get('Verwendungszweck','')}".lower()
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
        used_bank.add(bidx)
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

    matches_df = pd.DataFrame(matches).reset_index(drop=True)
    unmatched = bank_in.drop(index=list(used_bank), errors="ignore").copy().reset_index(drop=True)
    return matches_df, unmatched

# =========================================================
# SIDEBAR
# =========================================================
with st.sidebar:
    st.header("Import")
    fibu_file = st.file_uploader("1) Excel oder CSV", type=["xlsx", "xls", "csv"])
    bank_pdf = st.file_uploader("2) Bank PDF (optional)", type=["pdf"])
    st.divider()
    show_rows = st.number_input("Max Zeilen Tabellen", min_value=50, max_value=2000, value=300, step=50)
    st.divider()
    st.subheader("Schwellwerte")
    target_overdue_pct = st.slider("Ziel Überfällig Quote", min_value=0, max_value=50, value=10, step=1)
    target_dso = st.slider("Ziel DSO Tage", min_value=0, max_value=90, value=30, step=1)

if not fibu_file:
    st.info("Bitte Excel oder CSV hochladen.")
    st.stop()

# =========================================================
# LOAD SOURCE
# =========================================================
file_name = getattr(fibu_file, "name", "").lower()

if file_name.endswith(".csv"):
    df_norm = pd.read_csv(fibu_file, sep=None, engine="python", dtype="string")
    df_norm.columns = make_unique_columns([str(c).strip() for c in df_norm.columns])
    sheet_used = "CSV"
    header_row_used = 0
else:
    xls = pd.ExcelFile(fibu_file)
    with st.sidebar:
        sheet_used = st.selectbox("Excel Sheet", options=xls.sheet_names, index=0)
    raw = read_excel_raw(fibu_file, sheet_used)
    auto_header = auto_find_header_row(raw)
    with st.sidebar:
        header_row_used = st.number_input(
            "Header Zeile (0 basiert)",
            min_value=0,
            max_value=max(len(raw) - 1, 0),
            value=int(auto_header),
            step=1,
        )
    df_norm = normalize_with_header_row(raw, header_row_used)

df_norm = df_norm.reset_index(drop=True)
if df_norm.empty:
    st.error("Import Ergebnis ist leer. Prüfe Sheet und Header Zeile.")
    st.stop()

# =========================================================
# MAPPING
# =========================================================
cols = list(df_norm.columns)

auto_kunde = guess_col(cols, ["kunde", "debitor", "name"])
auto_re = guess_col(cols, ["re_n", "re-n", "re nr", "rechnungs", "beleg", "nummer"])
auto_redat = guess_col(cols, ["re_datum", "re-datum", "belegdat", "datum"])
auto_faellig = guess_col(cols, ["fällig", "faellig", "termin", "fälligkeit", "faelligkeit"])
auto_betrag = guess_col(cols, ["betrag (netto)", "betrag netto", "betrag (brutto)", "betrag brutto", "betrag", "brutto", "netto", "summe", "umsatz"])
auto_paid = guess_col(cols, ["gezahlt am", "gezahlt", "eingang", "ausgleich", "zahlung"])

with st.sidebar:
    st.divider()
    st.subheader("Mapping")
    col_kunde = st.selectbox("Kunde", options=["<leer>"] + cols, index=(["<leer>"] + cols).index(auto_kunde) if auto_kunde in cols else 0)
    col_re = st.selectbox("RE Nummer", options=["<leer>"] + cols, index=(["<leer>"] + cols).index(auto_re) if auto_re in cols else 0)
    col_redat = st.selectbox("Rechnungsdatum", options=cols, index=cols.index(auto_redat) if auto_redat in cols else 0)
    col_faellig = st.selectbox("Fälligkeit", options=["<leer>"] + cols, index=(["<leer>"] + cols).index(auto_faellig) if auto_faellig in cols else 0)
    col_betrag = st.selectbox("Betrag", options=cols, index=cols.index(auto_betrag) if auto_betrag in cols else 0)
    col_paid = st.selectbox("Zahldatum", options=["<keins>"] + cols, index=(["<keins>"] + cols).index(auto_paid) if auto_paid in cols else 0)

col_kunde = None if col_kunde == "<leer>" else col_kunde
col_re = None if col_re == "<leer>" else col_re
col_faellig = None if col_faellig == "<leer>" else col_faellig
col_paid = None if col_paid == "<keins>" else col_paid

inv = build_invoice_table(df_norm, col_kunde, col_re, col_redat, col_faellig, col_betrag, col_paid)
if inv.empty:
    st.error("Nach Mapping keine Datensätze übrig. Prüfe Rechnungsdatum und Betrag.")
    st.stop()

# =========================================================
# FILTERS
# =========================================================
inv["Monat"] = inv["RE_Datum"].dt.to_period("M").astype("string")
customers = sorted([c for c in inv["Kunde"].dropna().astype(str).unique().tolist() if c.strip()])

cF1, cF2, cF3 = st.columns([2, 2, 1])
with cF1:
    sel_customers = st.multiselect("Kunden Filter", options=customers, default=customers if customers else [])
with cF2:
    min_d = inv["RE_Datum"].min().date()
    max_d = inv["RE_Datum"].max().date()
    dr = st.date_input("Zeitraum", value=(min_d, max_d))
with cF3:
    compact_numbers = st.toggle("Kompaktzahlen", value=True)

d_from, d_to = (dr[0], dr[1]) if isinstance(dr, tuple) and len(dr) == 2 else (min_d, max_d)

f = inv.copy().reset_index(drop=True)
f = f[(f["RE_Datum"].dt.date >= d_from) & (f["RE_Datum"].dt.date <= d_to)].copy().reset_index(drop=True)
if sel_customers:
    f = f[f["Kunde"].isin(sel_customers)].copy().reset_index(drop=True)

today = pd.Timestamp(date.today())
f["Offen"] = f["Gezahlt_Am"].isna()
offen = f[f["Offen"]].copy().reset_index(drop=True)
bezahlt = f[~f["Offen"]].copy().reset_index(drop=True)

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

rev = float(f["Betrag"].sum())
op_sum = float(offen["Betrag"].sum()) if not offen.empty else 0.0
overdue_sum = float(offen.loc[offen["VerzugTage"] > 0, "Betrag"].sum()) if not offen.empty else 0.0
overdue_pct = (overdue_sum / op_sum * 100) if op_sum > 0 else 0.0

dso = None
if not bezahlt.empty:
    dso_val = (bezahlt["Gezahlt_Am"] - bezahlt["RE_Datum"]).dt.days.mean()
    dso = float(dso_val) if pd.notna(dso_val) else None

def forecast_sum(days: int) -> float:
    if offen.empty:
        return 0.0
    horizon = today + pd.Timedelta(days=days)
    tmp = offen[offen["Faellig"].notna()].copy()
    return float(tmp.loc[tmp["Faellig"] <= horizon, "Betrag"].sum())

cash_7 = forecast_sum(7)
cash_14 = forecast_sum(14)
cash_30 = forecast_sum(30)

missing_faellig = int(offen["Faellig"].isna().sum()) if not offen.empty else 0
missing_kunde = int((f["Kunde"].astype("string").str.strip() == "").sum())
missing_re = int((f["RE_Nr"].astype("string").str.strip() == "").sum())

dq_score = 100
dq_score -= min(30, missing_faellig)
dq_score -= min(20, missing_kunde)
dq_score -= min(20, missing_re)
dq_score = max(0, dq_score)

if overdue_pct <= target_overdue_pct:
    overdue_badge = badge_html("ok", f"Überfällig {overdue_pct:.1f}%")
elif overdue_pct <= target_overdue_pct + 10:
    overdue_badge = badge_html("warn", f"Überfällig {overdue_pct:.1f}%")
else:
    overdue_badge = badge_html("bad", f"Überfällig {overdue_pct:.1f}%")

if dso is None:
    dso_badge = badge_html("warn", "DSO n v")
elif dso <= target_dso:
    dso_badge = badge_html("ok", f"DSO {dso:.1f} T")
elif dso <= target_dso + 10:
    dso_badge = badge_html("warn", f"DSO {dso:.1f} T")
else:
    dso_badge = badge_html("bad", f"DSO {dso:.1f} T")

if dq_score >= 85:
    dq_badge = badge_html("ok", f"Datenqualität {dq_score}/100")
elif dq_score >= 70:
    dq_badge = badge_html("warn", f"Datenqualität {dq_score}/100")
else:
    dq_badge = badge_html("bad", f"Datenqualität {dq_score}/100")

def money_display(x: float) -> str:
    return eur_short(x) if compact_numbers else eur_full(x)

# =========================================================
# TOP BAR
# =========================================================
mandant_name = "Mandant"
if sel_customers and len(sel_customers) == 1 and sel_customers[0].strip():
    mandant_name = sel_customers[0].strip()
elif sel_customers and len(sel_customers) > 1:
    mandant_name = f"{len(sel_customers)} Kunden"

st.markdown(
    f"""
    <div class="topbar">
      <div>
        <div class="brand">Sohn Consult Executive Cash BI</div>
        <div class="sub">Quelle: {sheet_used} | Stand: {format_date_de(date.today())} | Zeitraum: {format_date_de(d_from)} bis {format_date_de(d_to)}</div>
      </div>
      <div class="pillRow">
        <div class="pill"><b>Scope:</b> {mandant_name}</div>
        <div class="pill"><b>Belege:</b> {len(f)}</div>
        <div class="pill"><b>Offen:</b> {len(offen)}</div>
        <div class="pill"><b>Bezahlt:</b> {len(bezahlt)}</div>
      </div>
    </div>
    """,
    unsafe_allow_html=True,
)

# =========================================================
# EXEC KPIS
# =========================================================
st.markdown("## Executive Überblick")
render_kpis(
    [
        {"label": "Umsatz im Zeitraum", "value": money_display(rev), "delta": "Basis: Import + Filter"},
        {"label": "Offene Posten", "value": money_display(op_sum), "delta": f"{overdue_badge}"},
        {"label": "Cash In Prognose 30 Tage", "value": money_display(cash_30), "delta": f"≤ 7T {money_display(cash_7)} | ≤ 14T {money_display(cash_14)}"},
        {"label": "Governance", "value": "OK" if overdue_pct <= target_overdue_pct else "Handlungsbedarf", "delta": f"{dso_badge} &nbsp; {dq_badge}"},
    ]
)

# =========================================================
# INSIGHTS
# =========================================================
def build_insights() -> list[str]:
    lines = []
    if op_sum <= 0:
        lines.append("Keine offenen Posten im Filter. Liquiditätsrisiko aus Forderungen ist aktuell niedrig.")
        return lines

    if overdue_pct > target_overdue_pct + 10:
        lines.append(f"Überfällig Quote {overdue_pct:.1f}%. Priorität: Top Debitoren eskalieren und Mahnprozess straffen.")
    elif overdue_pct > target_overdue_pct:
        lines.append(f"Überfällig Quote {overdue_pct:.1f}%. Empfohlen: Mahnlauf diese Woche.")
    else:
        lines.append(f"Überfällig Quote {overdue_pct:.1f}%. Forderungsmanagement im grünen Bereich.")

    if dso is None:
        lines.append("DSO nicht berechenbar, weil Zahldaten fehlen. Zahldatum Mapping prüfen.")
    elif dso > target_dso + 10:
        lines.append(f"DSO {dso:.1f} Tage. Empfehlung: Zahlungsbedingungen und Eskalationspfade überarbeiten.")
    elif dso > target_dso:
        lines.append(f"DSO {dso:.1f} Tage. Empfehlung: Top Debitoren aktiv steuern.")
    else:
        lines.append(f"DSO {dso:.1f} Tage. Zahlungsdisziplin im Zielkorridor.")

    by_c = f.groupby("Kunde", as_index=False)["Betrag"].sum().sort_values("Betrag", ascending=False)
    top3_share = (float(by_c["Betrag"].head(3).sum()) / rev * 100) if rev > 0 else 0.0
    if top3_share >= 60:
        lines.append(f"Klumpenrisiko hoch: Top 3 Anteil {top3_share:.1f}%. Zahlungsabsicherung prüfen.")
    elif top3_share >= 40:
        lines.append(f"Klumpenrisiko mittel: Top 3 Anteil {top3_share:.1f}%. Priorisierte Steuerung Top 3.")
    else:
        lines.append(f"Klumpenrisiko moderat: Top 3 Anteil {top3_share:.1f}%.")

    if missing_faellig > 0:
        lines.append(f"{missing_faellig} offene Posten ohne Fälligkeit. Forecast wird ungenauer. Datenpflege erforderlich.")
    if missing_kunde > 0:
        lines.append(f"{missing_kunde} Zeilen ohne Kunde. Debitorenfeld bereinigen.")
    if missing_re > 0:
        lines.append(f"{missing_re} Zeilen ohne RE Nummer. Nummernkreis herstellen, sonst schwächeres Matching.")

    lines.append("Nächste Aktionen: Top 5 überfällige Rechnungen priorisieren, Mahnlauf heute, Forecast Datenqualität erhöhen.")
    return lines[:8]

st.divider()
cI1, cI2 = st.columns([2, 1])

with cI1:
    st.markdown("### Consulting Insights")
    ins = build_insights()
    html = ["<div class='insights'>", "<div class='insTitle'>Empfehlungen auf einen Blick</div>"]
    for ln in ins:
        html.append(f"<div class='insLine'>• {ln}</div>")
    html.append(f"<div class='insMuted'>Hinweis: Aussagen basieren auf Import Mapping, Zeitraum Filter und Datenqualität (Score {dq_score}/100).</div>")
    html.append("</div>")
    st.markdown("\n".join(html), unsafe_allow_html=True)

with cI2:
    st.markdown("### Datenqualität")
    dq = pd.DataFrame(
        [
            {"Kriterium": "Offen ohne Fälligkeit", "Anzahl": missing_faellig},
            {"Kriterium": "Zeilen ohne Kunde", "Anzahl": missing_kunde},
            {"Kriterium": "Zeilen ohne RE Nummer", "Anzahl": missing_re},
        ]
    )
    st.dataframe(df_for_display(dq), width="stretch")

# =========================================================
# CFO CHARTS
# =========================================================
st.divider()
st.markdown("## CFO Charts")
cC1, cC2, cC3 = st.columns([2, 1, 1])

with cC1:
    st.markdown("### Cash In Forecast 60 Tage")
    if offen.empty or offen["Faellig"].isna().all():
        st.info("Forecast benötigt Fälligkeit. Aktuell fehlen Fälligkeiten oder es gibt keine offenen Posten.")
    else:
        tmp = offen[offen["Faellig"].notna()].copy()
        horizon_end = (today + pd.Timedelta(days=60)).date()
        tmp = tmp[tmp["Faellig"].dt.date <= horizon_end]
        tmp["Woche"] = tmp["Faellig"].dt.to_period("W").apply(lambda p: p.start_time.date())
        weekly = tmp.groupby("Woche", as_index=False)["Betrag"].sum().sort_values("Woche")
        weekly["Woche"] = pd.to_datetime(weekly["Woche"])
        fig = px.area(weekly, x="Woche", y="Betrag")
        fig.update_layout(margin=dict(l=10, r=10, t=10, b=10), height=320)
        fig.update_yaxes(tickformat=",.0f")
        st.plotly_chart(fig, width="stretch")

with cC2:
    st.markdown("### Aging Mix")
    if offen.empty:
        st.info("Keine offenen Posten.")
    else:
        aging = offen.groupby("Aging", as_index=False)["Betrag"].sum()
        fig = px.pie(aging, names="Aging", values="Betrag", hole=0.55)
        fig.update_layout(margin=dict(l=10, r=10, t=10, b=10), height=320, legend_title_text="")
        st.plotly_chart(fig, width="stretch")

with cC3:
    st.markdown("### Top Debitoren")
    if f.empty:
        st.info("Keine Daten im Filter.")
    else:
        by_c = f.groupby("Kunde", as_index=False)["Betrag"].sum().sort_values("Betrag", ascending=False)
        top = by_c.head(10).copy()
        fig = px.bar(top, x="Kunde", y="Betrag")
        fig.update_layout(margin=dict(l=10, r=10, t=10, b=10), height=320)
        fig.update_xaxes(tickangle=45)
        fig.update_yaxes(tickformat=",.0f")
        st.plotly_chart(fig, width="stretch")

# =========================================================
# DRILLDOWN
# =========================================================
st.divider()
st.markdown("## Drilldown")

if "drill" not in st.session_state:
    st.session_state["drill"] = "all"

d1, d2, d3, d4 = st.columns(4)
with d1:
    if st.button("Alle", use_container_width=True):
        st.session_state["drill"] = "all"
with d2:
    if st.button("Offen", use_container_width=True):
        st.session_state["drill"] = "open"
with d3:
    if st.button("Überfällig", use_container_width=True):
        st.session_state["drill"] = "overdue"
with d4:
    if st.button("> 60 Tage", use_container_width=True):
        st.session_state["drill"] = "gt60"

drill = st.session_state["drill"]
if drill == "all":
    drill_df = f.copy()
    title = "Alle Belege"
elif drill == "open":
    drill_df = offen.copy()
    title = "Offene Posten"
elif drill == "overdue":
    drill_df = offen.loc[offen["VerzugTage"] > 0].copy()
    title = "Überfällige Posten"
else:
    drill_df = offen.loc[offen["VerzugTage"] > 60].copy()
    title = "Überfällige Posten > 60 Tage"

drill_df = drill_df.reset_index(drop=True)
st.markdown(f"### {title}")

cols_show = [c for c in ["Kunde", "RE_Nr", "RE_Datum", "Faellig", "Betrag", "Gezahlt_Am", "VerzugTage", "Aging"] if c in drill_df.columns]
st.dataframe(
    df_for_display(drill_df[cols_show].head(300), money_cols=["Betrag"], date_cols=["RE_Datum", "Faellig", "Gezahlt_Am"]),
    width="stretch",
)

# =========================================================
# EXPORTS
# =========================================================
st.divider()
st.markdown("## Exporte")
e1, e2, e3 = st.columns([1, 1, 1])

with e1:
    st.download_button(
        "Export OP Liste Excel",
        data=to_excel_bytes(offen[cols_show], sheet_name="OP"),
        file_name="OP_Liste.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
with e2:
    by_cust = f.groupby("Kunde", as_index=False)["Betrag"].sum().sort_values("Betrag", ascending=False)
    st.download_button(
        "Export Kundenumsatz Excel",
        data=to_excel_bytes(by_cust, sheet_name="Kunden"),
        file_name="Kunden_Umsatz.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
with e3:
    st.download_button(
        "Export Belege Excel",
        data=to_excel_bytes(f[["Kunde", "RE_Nr", "RE_Datum", "Faellig", "Betrag", "Gezahlt_Am"]], sheet_name="Belege"),
        file_name="Belege.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

# =========================================================
# BANK SECTION
# =========================================================
st.divider()
st.markdown("## Bank Abgleich")

if not bank_pdf:
    st.info("Optional: Bank PDF hochladen, um Zahlungseingänge zuzuordnen.")
else:
    bank = parse_bank_pdf(bank_pdf)
    if bank.empty:
        st.warning("Keine Buchungen erkannt. Wenn das PDF ein Scan ist, braucht es OCR. Diese Version nutzt kein OCR.")
    else:
        matches, unmatched = reconcile_bank_vs_open(bank, offen)
        total_in = float(bank.loc[bank["Betrag"] > 0, "Betrag"].sum())
        match_rate = (len(matches) / max(1, len(bank[bank["Betrag"] > 0]))) * 100

        st.markdown("### Bank Summary")
        render_kpis(
            [
                {"label": "Bank Eingänge positiv", "value": eur_full(total_in), "delta": f"Match Quote {match_rate:.0f}%"},
                {"label": "Matches", "value": f"{len(matches)}", "delta": "Zuordnung nach RE Nummer oder Betrag"},
                {"label": "Unmatched", "value": f"{len(unmatched)}", "delta": "Manuelle Prüfung"},
                {"label": "Offene Posten", "value": eur_full(op_sum), "delta": "Basis OP"},
            ]
        )

        b1, b2 = st.columns([1, 1])
        with b1:
            st.markdown("### Matches")
            if matches.empty:
                st.info("Keine Matches gefunden.")
            else:
                mcols = [c for c in ["Buchung", "Betrag", "Gegenpartei", "RE_Nr", "Kunde", "Invoice_Betrag", "MatchType"] if c in matches.columns]
                st.dataframe(
                    df_for_display(matches[mcols].head(300), money_cols=["Betrag", "Invoice_Betrag"], date_cols=["Buchung"]),
                    width="stretch",
                )
        with b2:
            st.markdown("### Unmatched Eingänge")
            if unmatched.empty:
                st.success("Keine Unmatched Eingänge.")
            else:
                ucols = [c for c in ["Buchung", "Betrag", "Gegenpartei", "Verwendungszweck"] if c in unmatched.columns]
                st.dataframe(
                    df_for_display(unmatched[ucols].head(300), money_cols=["Betrag"], date_cols=["Buchung"]),
                    width="stretch",
                )

        x1, x2 = st.columns(2)
        with x1:
            st.download_button(
                "Export Bank Matches Excel",
                data=to_excel_bytes(matches, sheet_name="Matches"),
                file_name="Bank_Matches.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        with x2:
            st.download_button(
                "Export Bank Unmatched Excel",
                data=to_excel_bytes(unmatched, sheet_name="Unmatched"),
                file_name="Bank_Unmatched.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

# =========================================================
# DEBUG
# =========================================================
with st.expander("Import Debug Ansicht", expanded=False):
    st.write(f"Quelle: {sheet_used} | Header Zeile: {header_row_used} | Zeilen: {len(df_norm)} | Spalten: {len(df_norm.columns)}")
    st.dataframe(df_for_display(df_norm.head(50)), width="stretch")
