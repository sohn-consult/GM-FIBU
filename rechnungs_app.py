import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO

# 1. Konfiguration der Weboberfläche
st.set_page_config(page_title="Debitoren-Analyse 2025", layout="wide")

st.title("📊 Rechnungs-Auswertung & Export")
st.markdown("Analysiere deine Umsätze und exportiere die offenen Posten mit einem Klick.")

# Hilfsfunktion für den Excel-Export
def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Offene_Posten')
    processed_data = output.getvalue()
    return processed_data

# 2. Datei-Upload
uploaded_file = st.file_uploader("Bitte die Datei 'Debitoren 2025 neu.xlsx' hochladen", type=["xlsx"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file, sheet_name="2025 Debitoren")
        df.columns = [str(c).strip() for c in df.columns]

        col_datum = next((c for c in df.columns if "Datum" in c), None)
        col_betrag = next((c for c in df.columns if "Brutto" in c or "Betrag" in c), None)
        col_bezahlt = next((c for c in df.columns if "Bezahlt" in c or "Eingang" in c), None)

        if col_datum and col_betrag:
            df[col_datum] = pd.to_datetime(df[col_datum], errors='coerce')
            df = df.dropna(subset=[col_datum])
            df['Monat'] = df[col_datum].dt.strftime('%Y-%m')

            # --- ANALYSE ---
            monats_umsatz = df.groupby('Monat')[col_betrag].sum().reset_index()
            
            # Offene Posten Logik: Zeilen ohne Datum in der "Bezahlt"-Spalte
            if col_bezahlt:
                offene_posten = df[df[col_bezahlt].isna()].copy()
            else:
                offene_posten = pd.DataFrame()

            # --- WEB-AUSGABE ---
            kpi1, kpi2 = st.columns(2)
            kpi1.metric("Gesamtumsatz 2025", f"{df[col_betrag].sum():,.2f} €")
            kpi2.metric("Offene Beträge", f"{offene_posten[col_betrag].sum():,.2f} €", delta_color="inverse")

            # Diagramm
            st.subheader("📅 Monatliche Übersicht")
            fig = px.bar(monats_umsatz, x='Monat', y=col_betrag, color_discrete_sequence=['#007bff'])
            st.plotly_chart(fig, use_container_width=True)

            # Tabelle & Export
            st.subheader("🔴 Offene Rechnungen")
            if not offene_posten.empty:
                # Relevante Spalten für den Export auswählen
                cols_to_show = [col_datum, col_betrag]
                if "Kunde" in df.columns: cols_to_show.append("Kunde")
                if "Rechnungsnummer" in df.columns: cols_to_show.append("Rechnungsnummer")
                
                final_offene = offene_posten[cols_to_show].sort_values(by=col_datum)
                st.dataframe(final_offene, use_container_width=True)

                # --- EXPORT BUTTON ---
                excel_data = to_excel(final_offene)
                st.download_button(
                    label="📥 Offene Posten als Excel herunterladen",
                    data=excel_data,
                    file_name="Offene_Rechnungen_2025.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.success("Keine offenen Posten gefunden!")

    except Exception as e:
        st.error(f"Fehler: {e}")