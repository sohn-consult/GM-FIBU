import streamlit as st
import pandas as pd
import numpy as np
import re
from io import BytesIO
from datetime import date
import PyPDF2

# -----------------------------
# CONFIG
# -----------------------------
st.set_page_config(
    page_title="Sohn Consult | Liquidität & OP",
    page_icon="👔",
    layout="wide",
)

st.title("👔 Sohn Consult | Liquidität, Offene Posten, Bank Abgleich")
st.caption("Stabile Minimalversion für schnelle Beratungs Auswertungen (Excel + Bank PDF).")

# -----------------------------
# UTIL: safe display (Arrow-proof)
# -----------------------------
def df_for_display(df: pd.DataFrame) -> pd.DataFrame:
    """Make dataframe Arrow compatible for Streamlit display."""
    out = df.copy()
    for c in out.columns:
        if pd.api.types.is_datetime64_any_dtype(out[c]):
            out[c] = out[c].dt.strftime("%Y-%m-%d")
        elif pd.api.types.is_numeric_dtype(out[c]):
            # keep numeric
            pass
        else:
            out[c] = out[c].astype("string")
    return out


def format_eur(x) -> str:
    if x is None or (isinstance(x, float) and np.isnan(x)) or pd.isna(x):
        return "0,00 €"
    return f"{float(x):,.2f} €".replace(",", "X").replace(".", ",").replace("X", ".")


def to_excel_bytes(df: pd.DataFrame, sheet_name: str = "Export") -> bytes:
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    return bio.getvalue()


# -----------------------------
# UTIL: robust parsing
# -----------------------------
def _try_parse_date_series(s: pd.Series) -> pd.Series:
    """
    Robust date parsing without dayfirst surprises.
    Strategy:
      1) Try common German formats explicitly
      2) Fallback to pandas parser without dayfirst forcing
    """
    if s is None:
        return pd.to_datetime(pd.Series([pd.NaT] * 0))

    s2 = s.astype("string").str.strip()
    s2 = s2.replace({"": pd.NA, "nan": pd.NA, "NaT": pd.NA})

    # If already datetime-like, keep
    if pd.api.types.is_datetime64_any_dtype(s):
        return pd.to_datetime(s, errors="coerce")

    # 1) dd.mm.yyyy or dd.mm.yyyy hh:mm
    dt1 = pd.to_datetime(s2, format="%d.%m.%Y", errors="coerce")
    dt2 = pd.to_datetime(s2, format="%d.%m.%Y %H:%M", errors="coerce")
    dt3 = pd.to_datetime(s2, format="%d.%m.%Y %H:%M:%S", errors="coerce")

    # 2) yyyy-mm-dd (Excel often)
    dt4 = pd.to_datetime(s2, format="%Y-%m-%d", errors="coerce")
    dt5 = pd.to_datetime(s2, format="%Y-%m-%d %H:%M", errors="coerce")
    dt6 = pd.to_datetime(s2, format="%Y-%m-%d %H:%M:%S", errors="coerce")

    # combine: first non-NaT wins
    out = dt1
    for cand in [dt2, dt3, dt4, dt5, dt6]:
        out = out.fillna(cand)

    # fallback (no dayfirst forced)
    out = out.fillna(pd.to_datetime(s2, errors="coerce"))

    return out


def _to_number_series(s: pd.Series) -> pd.Series:
    """Convert EU formatted numbers safely to float."""
    if pd.api.types.is_numeric_dtype(s):
        return pd.to_numeric(s, errors="coerce")

    s2 = s.astype("string").str.strip()
    s2 = s2.str.replace("\u00a0", "", regex=False)  # nbsp
    s2 = s2.str.replace(".", "", regex=False)       # thousand sep
    s2 = s2.str.replace(",", ".", regex=False)      # decimal
    s2 = s2.str.replace("€", "", regex=False)
    return pd.to_numeric(s2, errors="coerce")


def find_header_row(raw: pd.DataFrame, required_keywords: list[str]) -> int | None:
    """
    Try to locate the row that contains the real header.
    Returns row index or None.
    """
    max_scan = min(len(raw), 50)
    for i in range(max_scan):
        row = raw.iloc[i].astype("string").str.lower().fillna("")
        hit = sum(any(k in cell for cell in row) for k in required_keywords for cell in row)
        # heuristic: at least 3 keyword hits
        if hit >= 3:
            return i
    return None


def load_invoice_excel(file) -> pd.DataFrame:
    """
    Load invoices from an uploaded Excel or CSV.
    Handles messy multi-row headers by detecting header row.
    Returns normalized dataframe with columns:
      Kunde, RE_Nr, RE_Datum, Faellig, Betrag, Gezahlt_Am
    """
    name = getattr(file, "name", "upload")
    if name.lower().endswith(".csv"):
        df = pd.read_csv(file, sep=None, engine="python", dtype="string")
        raw = df.copy()
        # assume header ok
        header_row = 0
        df.columns = [str(c).strip() for c in df.columns]
    else:
        # Read first sheet best-effort: often the main "Debitoren" sheet is first
        xls = pd.ExcelFile(file)
        sheet = xls.sheet_names[0]
        raw = pd.read_excel(file, sheet_name=sheet, header=None)
        header_row = find_header_row(raw, ["kunde", "re-nr", "re nr", "re-datum", "fällig", "faellig", "betrag"])
        if header_row is None:
            header_row = 0

        df = raw.iloc[header_row:].copy()
        df.columns = df.iloc[0].astype("string").str.strip()
        df = df.iloc[1:].copy()

    # Clean cols
    df.columns = [str(c).strip() for c in df.columns]
    df = df.loc[:, ~pd.Series(df.columns).astype(str).str.contains("^Unnamed", na=False)].copy()
    df.dropna(how="all", inplace=True)

    # Column mapping heuristics
    cols = [c for c in df.columns]

    def pick_col(keys: list[str]) -> str | None:
        for c in cols:
            cl = str(c).lower()
            if any(k in cl for k in keys):
                return c
        return None

    c_kunde = pick_col(["kunde", "debitor", "name"])
    c_re = pick_col(["re-nr", "re nr", "rechnungsnr", "rechnungsnummer", "beleg", "nummer"])
    c_redat = pick_col(["re-datum", "re datum", "belegdat", "datum"])
    c_faellig = pick_col(["fällig", "faellig", "termin", "fälligkeit", "faelligkeit"])
    c_betrag = pick_col(["betrag", "brutto", "netto", "umsatz", "summe"])
    c_paid = pick_col(["gezahlt", "ausgleich", "zahlung", "eingang"])

    # If your Debitoren sheet has "Betrag (netto)" / "Betrag (brutto)" and "gezahlt am"
    # prefer those if detected
    c_betrag_netto = pick_col(["betrag (netto)", "betrag netto"])
    c_betrag_brutto = pick_col(["betrag (brutto)", "betrag brutto"])
    c_paid_am = pick_col(["gezahlt am", "gezahlt_am", "payment date"])

    if c_betrag_netto is not None:
        c_betrag = c_betrag_netto
    elif c_betrag_brutto is not None:
        c_betrag = c_betrag_brutto

    if c_paid_am is not None:
        c_paid = c_paid_am

    # Hard guardrails: require at least date + amount
    if c_redat is None or c_betrag is None:
        raise ValueError(
            "Spalten nicht erkannt. Mindestens RE Datum und Betrag müssen existieren. "
            "Tipp: Prüfe, ob die Kopfzeile korrekt ist."
        )

    out = pd.DataFrame()
    out["Kunde"] = df[c_kunde].astype("string") if c_kunde else pd.Series([""] * len(df), dtype="string")
    out["RE_Nr"] = df[c_re].astype("string") if c_re else pd.Series([""] * len(df), dtype="string")
    out["RE_Datum"] = _try_parse_date_series(df[c_redat])
    out["Faellig"] = _try_parse_date_series(df[c_faellig]) if c_faellig else pd.NaT
    out["Betrag"] = _to_number_series(df[c_betrag])
    out["Gezahlt_Am"] = _try_parse_date_series(df[c_paid]) if c_paid else pd.NaT

    # Normalize
    out["Kunde"] = out["Kunde"].fillna("").astype("string").str.strip()
    out["RE_Nr"] = out["RE_Nr"].fillna("").astype("string").str.strip()

    out = out.dropna(subset=["RE_Datum", "Betrag"]).copy()
    out["Betrag"] = pd.to_numeric(out["Betrag"], errors="coerce").fillna(0.0)

    return out


# -----------------------------
# BANK PDF PARSER (Umsätze Druckansicht)
# -----------------------------
TX_RE = re.compile(
    r"\)\s*(\d{2}\.\d{2}\.\d{4})\s*(\d{2}\.\d{2}\.\d{4})\s*([+-]?\d{1,3}(?:\.\d{3})*,\d{2})",
    re.UNICODE,
)

DATE_IN_TEXT_RE = re.compile(r"\b(\d{2}\.\d{2}\.\d{4})\b")
INVOICE_NO_RE = re.compile(r"\b(20\d{6,}[-/]\S+|\d{6,})\b")  # broad, catches 20251230-RL... and long ids


def parse_bank_pdf(file) -> pd.DataFrame:
    """
    Parse bank statement PDF "Umsätze - Druckansicht" into transactions:
      Buchung, Wertstellung, Betrag, Gegenpartei, Verwendungszweck
    """
    reader = PyPDF2.PdfReader(file)
    lines: list[str] = []
    for p in reader.pages:
        t = p.extract_text() or ""
        for ln in t.splitlines():
            ln = ln.replace("\u00a0", " ").strip()
            if ln:
                lines.append(ln)

    # Filter some boilerplate
    def is_meta(ln: str) -> bool:
        l = ln.lower()
        return any(
            k in l
            for k in [
                "umsätze - druckansicht",
                "umsätze vom",
                "kontostand",
                "buchungwertstellung",
                "sichteinlagen",
                "iban",
                "de",
            ]
        )

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
            amt = _to_number_series(pd.Series([amt_raw.lstrip("+-")])).iloc[0]
            amt = float(amt) * sign if not pd.isna(amt) else np.nan

            tx.append(
                {
                    "Buchung": buchung,
                    "Wertstellung": wert,
                    "Betrag": amt,
                    "Gegenpartei": party.strip(),
                    "Verwendungszweck": " ".join(desc_buf).strip(),
                    "RawLine": ln,
                }
            )
            desc_buf = []
            continue

        # Heuristic for "counterparty line": relatively short, mostly letters
        # Keep last seen party until a transaction line finalizes it.
        if len(ln) <= 60 and not ln.startswith("(") and "DATUM" not in ln and "UHR" not in ln:
            # if it looks like a name line, update party and reset buffer
            # but keep buffer if it is clearly continuation
            if re.search(r"[A-Za-zÄÖÜäöüß]", ln):
                party = ln
                desc_buf = []
                continue

        # Otherwise treat as description continuation
        desc_buf.append(ln)

    df = pd.DataFrame(tx)
    if df.empty:
        return df

    df["Buchung"] = pd.to_datetime(df["Buchung"], errors="coerce")
    df["Wertstellung"] = pd.to_datetime(df["Wertstellung"], errors="coerce")
    df["Betrag"] = pd.to_numeric(df["Betrag"], errors="coerce")

    return df


def reconcile_bank_vs_invoices(bank: pd.DataFrame, inv_open: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Match bank incoming payments to open invoices.
    Matching priority:
      1) invoice number in bank text
      2) amount match within tolerance (0.02) and optional customer name hint
    Returns (matches, unmatched_bank_incoming)
    """
    if bank.empty or inv_open.empty:
        return pd.DataFrame(), bank

    bank_in = bank[bank["Betrag"] > 0].copy()
    inv = inv_open.copy()

    inv["MatchKey_RE"] = inv["RE_Nr"].astype("string").str.strip()
    inv["MatchKey_Kunde"] = inv["Kunde"].astype("string").str.lower().str.strip()

    used_inv_idx = set()
    matches = []

    # Pre index invoice numbers for quick lookup
    inv_map = {}
    for idx, r in inv.iterrows():
        key = str(r["MatchKey_RE"])
        if key and key != "nan":
            inv_map.setdefault(key, []).append(idx)

    # 1) Match by invoice number found in Verwendungszweck
    for bidx, br in bank_in.iterrows():
        text = f"{br.get('Gegenpartei','')} {br.get('Verwendungszweck','')}".strip()
        found = INVOICE_NO_RE.findall(text)
        chosen = None

        # try exact hits first
        for token in found:
            token = str(token).strip()
            if token in inv_map:
                # pick first unused invoice with that token
                for iidx in inv_map[token]:
                    if iidx not in used_inv_idx:
                        chosen = iidx
                        break
            if chosen is not None:
                break

        if chosen is not None:
            used_inv_idx.add(chosen)
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

    # remaining unmatched bank incoming
    matched_bank_idx = set()
    for m in matches:
        # find in bank by buchung+amount+gegenpartei (best effort)
        pass

    # Determine which bank rows were used: best effort by joining keys
    if matches:
        mdf = pd.DataFrame(matches)
        # build a key to mark used bank rows
        bank_in["_key"] = bank_in["Buchung"].astype("string") + "|" + bank_in["Betrag"].round(2).astype("string")
        mdf["_key"] = pd.to_datetime(mdf["Buchung"], errors="coerce").astype("string") + "|" + pd.to_numeric(mdf["Betrag"], errors="coerce").round(2).astype("string")
        used_keys = set(mdf["_key"].dropna().astype(str).tolist())
        bank_in_used = bank_in[bank_in["_key"].isin(used_keys)]
        matched_bank_idx = set(bank_in_used.index.tolist())
        bank_in = bank_in.drop(columns=["_key"], errors="ignore")

    # 2) Amount based matching for still open invoices and still unmatched bank rows
    tolerance = 0.02
    still_open = inv.drop(index=list(used_inv_idx), errors="ignore").copy()
    bank_rest = bank_in.drop(index=list(matched_bank_idx), errors="ignore").copy()

    if not still_open.empty and not bank_rest.empty:
        still_open["_amt"] = still_open["Betrag"].round(2)
        bank_rest["_amt"] = bank_rest["Betrag"].round(2)

        for bidx, br in bank_rest.iterrows():
            # candidates by amount
            cands = still_open[np.abs(still_open["_amt"] - br["_amt"]) <= tolerance]
            if cands.empty:
                continue

            # optional: prefer customer name hint in text
            text = f"{br.get('Gegenpartei','')} {br.get('Verwendungszweck','')}".lower()
            cands = cands.copy()
            cands["name_hit"] = cands["MatchKey_Kunde"].apply(lambda k: 1 if k and k != "nan" and k in text else 0)
            cands = cands.sort_values(["name_hit", "RE_Datum"], ascending=[False, True])

            chosen = None
            for iidx in cands.index.tolist():
                if iidx not in used_inv_idx:
                    chosen = iidx
                    break
            if chosen is None:
                continue

            used_inv_idx.add(chosen)
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

    # unmatched bank incoming after both rounds
    if matches_df.empty:
        unmatched_bank = bank[bank["Betrag"] > 0].copy()
    else:
        bank_in2 = bank[bank["Betrag"] > 0].copy()
        bank_in2["_key"] = bank_in2["Buchung"].astype("string") + "|" + bank_in2["Betrag"].round(2).astype("string")
        matches_df["_key"] = pd.to_datetime(matches_df["Buchung"], errors="coerce").astype("string") + "|" + pd.to_numeric(matches_df["Betrag"], errors="coerce").round(2).astype("string")
        used_keys = set(matches_df["_key"].dropna().astype(str).tolist())
        unmatched_bank = bank_in2[~bank_in2["_key"].isin(used_keys)].drop(columns=["_key"], errors="ignore")

    return matches_df.drop(columns=["_key"], errors="ignore"), unmatched_bank


# -----------------------------
# SIDEBAR: Upload
# -----------------------------
with st.sidebar:
    st.header("Import")
    fibu_file = st.file_uploader("1) Excel oder CSV (Debitoren, OP, Fibu)", type=["xlsx", "xls", "csv"])
    bank_pdf = st.file_uploader("2) Bank PDF (Umsätze Druckansicht)", type=["pdf"])

    st.divider()
    st.caption("Hinweis: Diese Version priorisiert Stabilität. Mapping ist automatisiert, Fokus: OP und Liquidität.")


# -----------------------------
# MAIN
# -----------------------------
if not fibu_file:
    st.info("Bitte zuerst eine Excel oder CSV hochladen.")
    st.stop()

try:
    inv = load_invoice_excel(fibu_file)
except Exception as e:
    st.error(f"Excel Import fehlgeschlagen: {e}")
    st.stop()

# Basic filters
inv["Monat"] = inv["RE_Datum"].dt.to_period("M").astype("string")
customers = sorted([c for c in inv["Kunde"].dropna().astype(str).unique().tolist() if c.strip()])

c1, c2, c3 = st.columns([2, 2, 1])
with c1:
    sel_customers = st.multiselect("Kunden Filter", options=customers, default=customers)
with c2:
    min_d = inv["RE_Datum"].min().date() if not inv.empty else date.today()
    max_d = inv["RE_Datum"].max().date() if not inv.empty else date.today()
    dr = st.date_input("Zeitraum", value=(min_d, max_d))
with c3:
    show_rows = st.number_input("Max Zeilen Tabellen", min_value=50, max_value=2000, value=300, step=50)

if isinstance(dr, tuple) and len(dr) == 2:
    d_from, d_to = dr
else:
    d_from, d_to = min_d, max_d

f = inv.copy()
f = f[(f["RE_Datum"].dt.date >= d_from) & (f["RE_Datum"].dt.date <= d_to)]
if sel_customers:
    f = f[f["Kunde"].isin(sel_customers)]

today = pd.Timestamp(date.today())
f["Offen"] = f["Gezahlt_Am"].isna()
offen = f[f["Offen"]].copy()
bezahlt = f[~f["Offen"]].copy()

# Aging
offen["VerzugTage"] = np.where(
    offen["Faellig"].notna(),
    (today - offen["Faellig"]).dt.days,
    np.nan,
)

def bucket(v):
    if pd.isna(v):
        return "Unbekannt"
    if v <= 0:
        return "Pünktlich"
    if v <= 30:
        return "1-30"
    if v <= 60:
        return "31-60"
    return ">60"

offen["Aging"] = offen["VerzugTage"].apply(bucket)

# KPIs
rev = f["Betrag"].sum()
op_sum = offen["Betrag"].sum()
overdue_sum = offen.loc[offen["VerzugTage"] > 0, "Betrag"].sum()

dso = 0.0
if not bezahlt.empty:
    dso = (bezahlt["Gezahlt_Am"] - bezahlt["RE_Datum"]).dt.days.mean()

k1, k2, k3, k4 = st.columns(4)
k1.metric("Umsatz im Zeitraum", format_eur(rev))
k2.metric("Offene Posten", format_eur(op_sum))
k3.metric("Überfällig", format_eur(overdue_sum))
k4.metric("Ø Zahlungsdauer DSO", f"{dso:.1f} Tage" if dso and dso > 0 else "N/A")

st.divider()

tab1, tab2, tab3 = st.tabs(["OP Überblick", "Kunden Fokus", "Bank Abgleich"])

with tab1:
    st.subheader("Offene Posten mit Aging")

    aging_sum = offen.groupby("Aging")["Betrag"].sum().reindex(["Pünktlich", "1-30", "31-60", ">60", "Unbekannt"]).fillna(0)
    a1, a2 = st.columns([1, 2])
    with a1:
        st.write("Aging Summen")
        st.dataframe(df_for_display(aging_sum.reset_index().rename(columns={"index": "Aging"})), width="stretch")
    with a2:
        show = offen.sort_values(["VerzugTage", "Faellig"], ascending=[False, True]).copy()
        show_cols = ["Kunde", "RE_Nr", "RE_Datum", "Faellig", "Betrag", "VerzugTage", "Aging"]
        st.dataframe(df_for_display(show[show_cols].head(int(show_rows))), width="stretch")

    st.download_button(
        "Excel Export OP Liste",
        data=to_excel_bytes(offen[["Kunde", "RE_Nr", "RE_Datum", "Faellig", "Betrag", "VerzugTage", "Aging"]]),
        file_name="OP_Liste.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

with tab2:
    st.subheader("Kunden Fokus: Top Debitoren und Klumpen")

    by_cust = f.groupby("Kunde")["Betrag"].sum().sort_values(ascending=False).reset_index()
    top = by_cust.head(15).copy()
    st.dataframe(df_for_display(top), width="stretch")

    top3_share = (by_cust["Betrag"].head(3).sum() / rev * 100) if rev > 0 else 0
    st.metric("Klumpenrisiko Top 3", f"{top3_share:.1f}%")

    st.download_button(
        "Excel Export Kundenumsatz",
        data=to_excel_bytes(by_cust, sheet_name="Kunden"),
        file_name="Kunden_Umsatz.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

with tab3:
    st.subheader("Bank Abgleich (PDF Umsätze Druckansicht)")

    if not bank_pdf:
        st.info("Bitte Bank PDF hochladen. Erwartet wird das Format wie in deiner Unterkonto PDF (Umsätze Druckansicht).")
        st.stop()

    try:
        bank = parse_bank_pdf(bank_pdf)
    except Exception as e:
        st.error(f"Bank PDF Parsing fehlgeschlagen: {e}")
        st.stop()

    if bank.empty:
        st.warning("Keine Buchungen erkannt. Wenn das PDF gescannt ist (Bild PDF), braucht es OCR. Diese Version nutzt kein OCR.")
        st.stop()

    st.write("Erkannte Buchungen (Auszug)")
    st.dataframe(df_for_display(bank[["Buchung", "Wertstellung", "Betrag", "Gegenpartei", "Verwendungszweck"]].head(80)), width="stretch")

    # Reconcile against open invoices
    open_invoices = offen.copy()
    matches, unmatched = reconcile_bank_vs_invoices(bank, open_invoices)

    st.divider()
    m1, m2, m3 = st.columns(3)
    m1.metric("Bank Eingänge gesamt", format_eur(bank.loc[bank["Betrag"] > 0, "Betrag"].sum()))
    m2.metric("Matches", f"{len(matches)}")
    m3.metric("Unmatched Eingänge", f"{len(unmatched)}")

    st.write("Matches (Bank ↔ OP)")
    if matches.empty:
        st.info("Noch keine Matches gefunden. Typische Gründe: keine RE Nummer im Verwendungszweck oder Beträge weichen ab.")
    else:
        show_m = matches[["Buchung", "Betrag", "Gegenpartei", "RE_Nr", "Kunde", "Invoice_Betrag", "MatchType"]].copy()
        st.dataframe(df_for_display(show_m.head(int(show_rows))), width="stretch")

    st.write("Unmatched Bank Eingänge (prüfen)")
    if unmatched.empty:
        st.success("Keine offenen Bank Eingänge ohne Zuordnung.")
    else:
        st.dataframe(df_for_display(unmatched[["Buchung", "Betrag", "Gegenpartei", "Verwendungszweck"]].head(int(show_rows))), width="stretch")

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
            "Excel Export Unmatched",
            data=to_excel_bytes(unmatched, sheet_name="Unmatched"),
            file_name="Bank_Unmatched.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
