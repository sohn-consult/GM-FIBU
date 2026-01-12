import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO

# 1. Grundkonfiguration
st.set_page_config(page_title="Sohn-Consult Auswertung Fibu", layout="wide")

# Styling für ein professionelles Auftreten
st.markdown("""
    <style>
    .main { background-color: #f8f9fa; }
    .stMetric { background-color: #ffffff; padding: 15px; border-radius: 10px; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }
    </style>
    """, unsafe_allow_html=True)

st.title("📊 Sohn-Consult Auswertung Fibu")
st.subheader("Finanzübersicht & Debitoren-Analyse 2025")

# 2. Datei-Upload
uploaded_file = st.file_uploader("Excel-Datei (Debitoren 2025 neu.xlsx) hier hochladen", type=["xlsx"])

def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Offene_Posten')
    return output.getvalue()

if uploaded_file:
    try:
        # Wir laden das Blatt "2025 Debitoren"
        df = pd.read_excel(uploaded_file, sheet_name="2025 Debitoren")
        
        # Spaltennamen bereinigen
        df.columns = [str(c).strip() for c in df.columns]

        # Automatische Spaltenerkennung (Sucht nach Schlüsselwörtern)
        col_datum = next((c for c in df.columns if "Datum" in c), None)
        col_betrag = next((c for c in df.columns if "Brutto" in c or "Summe" in c), None)
        col_bezahlt = next((c for c in df.columns if "Bezahlt" in c or "Eingang" in c), None)
        col_kunde = next((c for c in df.columns if "Kunde" in c or "Name" in c), "Kunde")

        if col_datum and col_betrag:
            # Datenkonvertierung
            df[col_datum] = pd.to_datetime(df[col_datum], errors='coerce')
            df = df.dropna(subset=[col_datum])
            
            # Beträge in Zahlen umwandeln (falls nötig)
            if df[col_betrag].dtype == 'object':
                df[col_betrag] = pd.to_numeric(df[col_betrag].astype(str).str.replace('.', '').str.replace(',', '.'), errors='coerce')
            
            df['Monat'] = df[col_datum].dt.strftime('%Y-%m')

            # --- BERECHNUNGEN ---
            # 1. Monatlicher Umsatz (geschriebene Rechnungen)
            monats_umsatz = df.groupby('Monat')[col_betrag].sum().reset_index()

            # 2. Offene Posten (Wenn in 'Bezahlt am' nichts steht)
            # Wir prüfen auf leere Zellen (NaN)
            offene_posten = df[df[col_bezahlt].isna()].copy()
            
            # --- WEB-AUSGABE (DASHBOARD) ---
            
            # KPI-Reihe
            m1, m2, m3 = st.columns(3)
            gesamtsumme = df[col_betrag].sum()
            offen_summe = offene_posten[col_betrag].sum()
            quote = (offen_summe / gesamtsumme * 100) if gesamtsumme > 0 else 0

            m1.metric("Gesamt Fakturiert", f"{gesamtsumme:,.2f} €")
            m2.metric("Offene Forderungen", f"{offen_summe:,.2f} €", delta=f"{quote:.1f}% Quote", delta_color="inverse")
            m3.metric("Anzahl offene Rechnungen", f"{len(offene_posten)}")

            st.divider()

            # Grafik-Sektion
            col_left, col_right = st.columns([2, 1])

            with col_left:
                st.subheader("📈 Umsatzentwicklung pro Monat")
                fig = px.bar(monats_umsatz, x='Monat', y=col_betrag, 
                             text_auto='.2s',
                             color_discrete_sequence=['#1f77b4'],
                             labels={col_betrag: "Umsatz in €", "Monat": "Monat"})
                st.plotly_chart(fig, use_container_width=True)

            with col_right:
                st.subheader("ℹ️ Analyse-Info")
                st.info(f"""
                **Gefundene Spalten:**
                - Datum: `{col_datum}`
                - Betrag: `{col_betrag}`
                - Zahlungsstatus: `{col_bezahlt}`
                
                Die Berechnung der offenen Posten basiert auf leeren Einträgen in der Spalte `{col_bezahlt}`.
                """)

            # Tabellen-Sektion
            st.divider()
            st.subheader("⚠️ Offene Posten (Nicht bezahlt)")
            
            if not offene_posten.empty:
                # Wichtige Spalten für die Anzeige auswählen
                display_cols = [c for c in [col_datum, col_kunde, col_betrag, "Beleg-Nr."] if c in df.columns]
                final_table = offene_posten[display_cols].sort_values(by=col_datum)
                
                st.dataframe(final_table, use_container_width=True)

                # Export Button
                st.download_button(
                    label="📥 Liste als Excel exportieren",
                    data=to_excel(final_table),
                    file_name="Sohn_Consult_Offene_Posten.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.success("Hervorragend! Alle Rechnungen sind als bezahlt markiert.")

        else:
            st.error(f"Spalten nicht erkannt. In deiner Tabelle wurden gefunden: {list(df.columns)}")

    except Exception as e:
        st.error(f"Fehler beim Lesen des Tabellenblatts '2025 Debitoren': {e}")
        st.info("Bitte stelle sicher, dass die Excel-Datei ein Blatt mit dem exakten Namen '2025 Debitoren' enthält.")
