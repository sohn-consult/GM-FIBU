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
    st.sidebar.header("⚙️ Einstellungen")
    
    # Schritt 1: Header-Zeile bestimmen
    header_idx = st.sidebar.number_input("In welcher Zeile stehen die Spaltennamen? (z.B. 1, 2, 3...)", 
                                        min_value=1, value=3, step=1)
    
    try:
        # Laden der Excel-Datei
        xl = pd.ExcelFile(uploaded_file)
        sheet_name = st.sidebar.selectbox("Tabellenblatt wählen", xl.sheet_names, 
                                         index=0 if "2025 Debitoren" not in xl.sheet_names else xl.sheet_names.index("2025 Debitoren"))
        
        # Einlesen mit der gewählten Header-Zeile (header_idx - 1 weil Python bei 0 anfängt zu zählen)
        df = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=header_idx-1)
        
        # Leere Spalten/Zeilen entfernen
        df = df.loc[:, ~df.columns.str.contains('^Unnamed')].dropna(how='all', axis=0)

        st.sidebar.markdown("---")
        st.sidebar.subheader("📍 Spalten zuordnen")
        
        # Spalten-Mapping durch den Nutzer
        all_cols = df.columns.tolist()
        
        col_datum = st.sidebar.selectbox("Spalte für RE-Datum", all_cols, index=0 if len(all_cols)>0 else 0)
        col_betrag = st.sidebar.selectbox("Spalte für Brutto-Betrag", all_cols, index=1 if len(all_cols)>1 else 0)
        col_bezahlt = st.sidebar.selectbox("Spalte für 'gezahlt am' (Zahlungseingang)", all_cols, index=2 if len(all_cols)>2 else 0)
        col_kunde = st.sidebar.selectbox("Spalte für Kunde/Projekt", all_cols, index=0)

        if st.sidebar.button("Analyse starten"):
            # --- DATENVERARBEITUNG ---
            # Kopie erstellen um Warnungen zu vermeiden
            data = df.copy()
            
            # Datum konvertieren
            data[col_datum] = pd.to_datetime(data[col_datum], errors='coerce')
            data = data.dropna(subset=[col_datum])
            
            # Betrag konvertieren
            if data[col_betrag].dtype == 'object':
                data[col_betrag] = pd.to_numeric(data[col_betrag].astype(str).str.replace('.', '').str.replace(',', '.'), errors='coerce')
            
            data['Monat'] = data[col_datum].dt.strftime('%Y-%m')

            # --- AUSWERTUNG ---
            monats_umsatz = data.groupby('Monat')[col_betrag].sum().reset_index()
            offene_posten = data[data[col_bezahlt].isna()].copy()

            # --- UI AUSGABE ---
            kpi1, kpi2, kpi3 = st.columns(3)
            kpi1.metric("Gesamtumsatz", f"{data[col_betrag].sum():,.2f} €")
            kpi2.metric("Offen", f"{offene_posten[col_betrag].sum():,.2f} €", delta_color="inverse")
            kpi3.metric("Anzahl Rechnungen", len(data))

            st.markdown("### 📈 Monatliche Rechnungsstellung")
            fig = px.bar(monats_umsatz, x='Monat', y=col_betrag, text_auto='.2s', color_discrete_sequence=['#1f77b4'])
            st.plotly_chart(fig, use_container_width=True)

            st.markdown("### 🔴 Offene Posten")
            if not offene_posten.empty:
                display_cols = [col_datum, col_kunde, col_betrag]
                st.dataframe(offene_posten[display_cols].sort_values(by=col_datum), use_container_width=True)
                
                # Export
                st.download_button(
                    label="📥 Diese Liste als Excel exportieren",
                    data=to_excel(offene_posten[display_cols]),
                    file_name="Sohn_Consult_Offene_Posten.xlsx"
                )
            else:
                st.success("Alle Rechnungen sind bezahlt!")
        else:
            st.info("Bitte nimm die Einstellungen in der linken Seitenleiste vor und klicke auf 'Analyse starten'.")

    except Exception as e:
        st.error(f"Fehler beim Einlesen: {e}")
        st.info("Tipp: Erhöhe die Zahl der 'Header-Zeile' in der Sidebar, bis die richtigen Spaltennamen erscheinen.")
