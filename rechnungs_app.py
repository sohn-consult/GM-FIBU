import streamlit as st
import pandas as pd
import numpy as np
import re
import textwrap
from io import BytesIO, StringIO
from datetime import date
import plotly.express as px

# Optional: Bank PDF Parsing (nur wenn installiert und PDF textbasiert ist)
try:
    import PyPDF2
    HAS_PDF = True
except Exception:
    HAS_PDF = False

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
      min-height: 86px;
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
def existing_cols(df: pd.DataFrame, wanted: list[str]) -> list[str]:
    return [c for c in wanted if c in df.columns]

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

def money_display(x: float, compact: bool) -> str:
    return eur_short(x) if compact else eur_full(x)

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
    # WICHTIG: kein eingerücktes HTML, sonst rendert Streamlit das als Codeblock
    cards = []
    for it in items:
        cards.append(
            textwrap.dedent(f"""
            <div class="kpiCard">
              <div class="kpiStripe"></div>
              <div class="kpiLabel">{it.get("label","")}</div>
              <div class="kpiValue">{it.get("value","")}</div>
              <div class="kpiDelta">{it.get("delta","")}</div>
            </div>
            """).strip()
        )
    html = textwrap.dedent(f"""
    <div class="kpiGrid">
      {''.join(cards)}
    </div>
    """).strip()
    st.markdown(html, unsafe_allow_html=True)

# =========================================================
# PARSERS
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

    def conv(val):
        if val is None:
            return np.nan
        v = str(val).strip()
        if v == "" or v.lower() == "nan":
            return np.nan

        if "," in v and "." in v:
            if v.rfind(",") > v.rfind("."):
                v = v.replace(".", "").replace(",", ".")
            else:
                v = v.replace(",", "")
            try:
                return float(v)
            except Exception:
                return np.nan

        if "," in v and "." not in v:
            v = v.replace(".", "").replace(",", ".")
            try:
                return float(v)
            except Exception:
                return np.nan

        if "." in v and "," not in v:
            try:
                return float(v)
            except Exception:
                return np.nan

        try:
            return float(v)
        except Exception:
            return np.nan

    return x.apply(conv)

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
# EXCEL LOADING
# =========================================================
def read_excel_raw(file, sheet_name: str) -> pd.DataFrame:
    return pd.read_excel(file, sheet_name=sheet_name, header=None, engine="openpyxl")

def auto_find_header_row(raw: pd.DataFrame) -> int:
    keywords = ["kunde", "debitor", "name", "rechnung", "re", "datum", "fällig", "faellig", "betrag", "brutto", "netto", "gezahlt", "eingang"]
    best_i = 0
    best_score = -1
    scan = min(len(raw), 60)
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
INVOICE_NO_RE = re.compile(r"\b(20\d{6,}[-/]\S+|\d{6,})\b")
TX_RE = re.compile(r"\)\s*(\d{2}\.\d{2}\.\d{4})\s*(\d{2}\.\d{2}\.\d{4})\s*([+-]?\d{1,3}(?:\.\d{3})*,\d{2})")

def parse_bank_pdf(file) -> pd.DataFrame:
    if not HAS_PDF:
        return pd.DataFrame()

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
    bank_pdf = st.file_uploader("2) Bank PDF optional", type=["pdf"])
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
# LOAD
# =========================================================
file_name = getattr(fibu_file, "name", "").lower()

if file_name.endswith(".csv"):
    content = fibu_file.getvalue().decode("utf-8", errors="ignore")
    df_norm = pd.read_csv(StringIO(content), sep=None, engine="python", dtype="string")
    df_norm.columns = make_unique_columns([str(c).strip() for c in df_norm.columns])
    sheet_used = "CSV"
else:
    xls = pd.ExcelFile(fibu_file)
    with st.sidebar:
        sheet_used = st.selectbox("Excel Sheet", options=xls.sheet_names, index=0)
    raw = read_excel_raw(fibu_file, sheet_used)
    auto_header = auto_find_header_row(raw)
    with st.sidebar:
        header_row_used = st.number_input(
            "Header Zeile 0 basiert",
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
auto_betrag = guess_col(cols, ["betrag (brutto)", "betrag brutto", "betrag (netto)", "betrag netto", "betrag", "brutto", "netto", "summe", "umsatz"])
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
f["VerzugTage"] = np.where(f["Faellig"].notna(), (today - f["Faellig"]).dt.days, np.nan)

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

f["Aging"] = f["VerzugTage"].apply(aging_bucket)

offen = f[f["Offen"]].copy().reset_index(drop=True)
bezahlt = f[~f["Offen"]].copy().reset_index(drop=True)

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
dq_score -= min(20,
