import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO

# 1. Grundkonfiguration
st.set_page_config(page_title="Sohn-Consult Auswertung Fibu", layout="wide")

st.title("📊 Sohn-Consult Auswertung Fibu")
st.markdown("---")

# Hilfsfunktion für den Excel-Export
def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Offene_Posten')
    return output.getvalue()

# 2. Datei-Upload
uploaded_file = st.file_uploader("Excel-Datei (Debitoren 2025 neu.xlsx) hochladen", type=["xlsx"])

if uploaded_file:
    # Sidebar für Einstellungen
    st.sidebar.header("⚙️ Einstellungen & Filter")
    
    # Header-Zeile bestimmen
    header_idx = st.sidebar.number_input("In welcher Zeile stehen die Spaltennamen?", 
                                        min_value=1, value=3, step=1)
    
    try:
        # Laden der Excel-Datei
        xl = pd.ExcelFile(uploaded_file)
        sheet_name = st.sidebar.selectbox("Tabellenblatt wählen", xl.sheet_names, 
                                         index=0 if "2025" in str(xl.sheet_names) else 0)
        
        df = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=header_idx-1)
        
        # Leere Spalten entfernen
        df = df.loc[:, ~df.columns.str.contains('^Unnamed')].dropna(how='all', axis=0)

        st.sidebar.markdown("---")
        st.sidebar.subheader("📍 Spalten zuordnen")
        
        all_cols = df.columns.tolist()
        col_datum = st.sidebar.selectbox("Rechnungsdatum", all_cols, index=2 if len(all_cols)>2 else 0)
        col_betrag = st.sidebar.selectbox("Rechnungsbetrag (Brutto)", all_cols, index=16 if len(all_cols)>16 else 0)
        col_bezahlt = st.sidebar.selectbox("Zahlungseingang (Spalte)", all_cols, index=17 if len(all_cols)>17 else 0)
        col_kunde = st.sidebar.selectbox("Kunde/Projekt", all_cols, index=0)

        # --- DATEN-VORBEREITUNG ---
        data = df.copy()
        data[col_datum] = pd.to_datetime(data[col_datum], errors='coerce')
        data = data.dropna(subset=[col_datum]) 
        
        if data[col_betrag].dtype == 'object':
            data[col_betrag] = pd.to_numeric(data[col_betrag].astype(str).str.replace('.', '').str.replace(',', '.'), errors='coerce')

        # --- FILTER-BEREICH IN DER SIDEBAR ---
        st.sidebar.markdown("---")
        st.sidebar.subheader("🔍 Filter")

        # 1. Kunden-Filter (Mehrfachauswahl)
        alle_kunden = sorted(data[col_kunde].dropna().unique().tolist())
        selected_kunden = st.sidebar.multiselect("Kunden auswählen", options=alle_kunden, default=alle_kunden)

        # 2. Zeitraum-Filter
        min_date = data[col_datum].min().date()
        max_date = data[col_datum].max().date()
        selected_dates = st.sidebar.date_input("Zeitraum wählen", [min_date, max_date])

        # Anwendung der Filter
        if isinstance(selected_dates, list) and len(selected_dates) == 2:
            start_date, end_date = selected_dates
            
            # Filtern nach Datum UND Kunde
            filtered_data = data[
                (data[col_datum].dt.date >= start_date) & 
                (data[col_datum].dt.date <= end_date) &
                (data[col_kunde].isin(selected_kunden))
            ].copy()

            # --- AUSWERTUNG ---
            offene_posten = filtered_data[filtered_data[col_bezahlt].isna()].copy()
            filtered_data['Monat'] = filtered_data[col_datum].dt.strftime('%Y-%m')
            monats_umsatz = filtered_data.groupby('Monat')[col_betrag].sum().reset_index()

            # --- DASHBOARD AUSGABE ---
            kpi1, kpi2, kpi3 = st.columns(3)
            total_sum = filtered_data[col_betrag].sum()
            open_sum = offene_posten[col_betrag].sum()
            
            kpi1.metric("Umsatz (gefiltert)", f"{total_sum:,.2f} €")
            kpi2.metric("Offen (gefiltert)", f"{open_sum:,.2f} €", delta=f"{(open_sum/total_sum*100):.1f}%" if total_sum > 0 else "0%")
            kpi3.metric("Anzahl Belege", len(filtered_data))

            st.markdown(f"### 📈 Monatliche Rechnungsstellung")
            if not monats_umsatz.empty:
                fig = px.bar(monats_umsatz, x='Monat', y=col_betrag, text_auto='.2s', 
                             color_discrete_sequence=['#1f77b4'], labels={col_betrag: "Summe Brutto"})
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.warning("Keine Daten für die gewählte Kombination vorhanden.")

            # Tabellen-Ansicht
            tab1, tab2 = st.tabs(["🔴 Offene Posten (gefiltert)", "✅ Alle Buchungen im Filter"])
            
            with tab1:
                if not offene_posten.empty:
                    display_cols = [col_datum, col_kunde, col_betrag]
                    st.dataframe(offene_posten[display_cols].sort_values(by=col_datum), use_container_width=True)
                    
                    st.download_button(
                        label="📥 Diese Auswahl exportieren",
                        data=to_excel(offene_posten[display_cols]),
                        file_name=f"Sohn_Consult_Export.xlsx"
                    )
                else:
                    st.success("Keine offenen Posten für diesen Filter!")

            with tab2:
                st.dataframe(filtered_data.sort_values(by=col_datum), use_container_width=True)

        else:
            st.info("Bitte wähle im Kalender links den Zeitraum fertig aus.")

    except Exception as e:
        st.error(f"Verarbeitungsfehler: {e}")
