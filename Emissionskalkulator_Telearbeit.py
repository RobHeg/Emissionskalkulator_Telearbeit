#Emissionskalkulator_Telearbeit.py
#setup: pip install streamlit gspread oauth2client pandas plotly
#run with:
#streamlit run Emissionskalkulator_Telearbeit.py
#datasheet is here: https://docs.google.com/spreadsheets/d/1YbvSFcHNWEOmGbXMenB-kHTXnDG0oEJ8qTTB0lB_zyQ/edit?usp=sharing
#use credentials.json in same folder as Emissionskalkulator_Telearbeit.py, or create own google API and share datasheet with service account

import streamlit as st
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
import plotly.express as px

def update_values(): 
    arbeitsweg_wert = st.session_state.arbeitsweg_wert 
    induzierte_wege_wert = st.session_state.induzierte_wege_wert 
    praesenzarbeitsort_wert = st.session_state.praesenzarbeitsort_wert 
    telearbeitsort_wert = st.session_state.telearbeitsort_wert 
    shared_desk_wert = st.session_state.shared_desk_wert 
    return [arbeitsweg_wert, induzierte_wege_wert, praesenzarbeitsort_wert, telearbeitsort_wert, shared_desk_wert]

# Google Sheets API setup
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
client = gspread.authorize(creds)

# Open the Google Sheet - must be shared with 
gsheet = client.open("Emissionskalkulator_Telearbeit")
ws_default_Werte = gsheet.worksheet("default_Werte")
ws_Arbeitsweg = gsheet.worksheet("Arbeitsweg")
ws_ppg = gsheet.worksheet("Primärenergiebedarf_Präsenzarbeitsort_Gebäude")
ws_pph = gsheet.worksheet("Primärenergiebedarf_Präsenzarbeitsort_Heizen")
ws_pth = gsheet.worksheet("Primärenergiebedarf_Telearbeitsort_Heizen")

#default_Werte data preparation
data_default_Werte = ws_default_Werte.get_all_records()
df_default_Werte = pd.DataFrame(data_default_Werte)

categories = ['Arbeitsweg', 'induzierte Wege', 'Primärenergiebedarf Präsenzarbeitsort', 'Primärenergiebedarf Telearbeitsort', 'Shared-Desk-Lösungen']

if 'anz_telearbeitstage' not in st.session_state:
    st.session_state.anz_telearbeitstage = 1
anz_telearbeitstage = st.session_state.anz_telearbeitstage

if 'arbeitsweg_wert_day' not in st.session_state:
    st.session_state.arbeitsweg_wert_day = df_default_Werte.loc[df_default_Werte['Kategorie'] == 'Arbeitsweg', 'Änderungen_Treibhauspotenzial_je_Telearbeitstag'].values[0]
if 'induzierte_wege_wert_day' not in st.session_state:
    st.session_state.induzierte_wege_wert_day = df_default_Werte.loc[df_default_Werte['Kategorie'] == 'induzierte Wege', 'Änderungen_Treibhauspotenzial_je_Telearbeitstag'].values[0]
if 'praesenzarbeitsort_wert_day' not in st.session_state:
    st.session_state.praesenzarbeitsort_wert_day = df_default_Werte.loc[df_default_Werte['Kategorie'] == 'Primärenergiebedarf Präsenzarbeitsort', 'Änderungen_Treibhauspotenzial_je_Telearbeitstag'].values[0]
if 'telearbeitsort_wert_day' not in st.session_state:
    st.session_state.telearbeitsort_wert_day = df_default_Werte.loc[df_default_Werte['Kategorie'] == 'Primärenergiebedarf Telearbeitsort', 'Änderungen_Treibhauspotenzial_je_Telearbeitstag'].values[0]
if 'shared_desk_wert_day' not in st.session_state:
    st.session_state.shared_desk_wert_day = df_default_Werte.loc[df_default_Werte['Kategorie'] == 'Shared-Desk-Lösungen', 'Änderungen_Treibhauspotenzial_je_Telearbeitstag'].values[0]
if 'induzierte_wege_wert_anteil' not in st.session_state:
    st.session_state.induzierte_wege_wert_anteil =  st.session_state.induzierte_wege_wert_day * 0.62

if 'arbeitsweg_wert' not in st.session_state:
    st.session_state.arbeitsweg_wert = st.session_state.arbeitsweg_wert_day * anz_telearbeitstage
if 'induzierte_wege_wert' not in st.session_state:
    st.session_state.induzierte_wege_wert =  st.session_state.induzierte_wege_wert_day * anz_telearbeitstage
if 'praesenzarbeitsort_wert' not in st.session_state:
    st.session_state.praesenzarbeitsort_wert = st.session_state.praesenzarbeitsort_wert_day * anz_telearbeitstage
if 'telearbeitsort_wert' not in st.session_state:
    st.session_state.telearbeitsort_wert = st.session_state.telearbeitsort_wert_day * anz_telearbeitstage
if 'shared_desk_wert' not in st.session_state:
    st.session_state.shared_desk_wert = st.session_state.shared_desk_wert_day * anz_telearbeitstage


values = update_values()






######################################################################################################################################
# begin layout
######################################################################################################################################


st.title("Emissionskalkulator Telearbeit")

st.write(f"Schätzen Sie mit Hilfe dieses prototypischen Kalkulators ab, ob sich ein Telearbeitsmodell gegenüber der vollständigen Präsenzarbeit in Ihrem Fall lohnt! Sie können dazu in den folgenden Kategorien (drop-down-Bereiche) Angaben zu Ihrem individuellen Fall machen. Falls Sie keine genauere Angaben tätigen, werden Durchschnittswerte genommen.")

######################################################################################################################################
# Anzahl Telearbeitstage

st.session_state.anz_telearbeitstage = st.slider("An wie vielen Tagen in der Woche führen Sie Telearbeit aus?", min_value=1, max_value=5, value=1)
st.session_state.arbeitsweg_wert = st.session_state.arbeitsweg_wert_day * st.session_state.anz_telearbeitstage
st.session_state.induzierte_wege_wert = st.session_state.induzierte_wege_wert_day * st.session_state.anz_telearbeitstage
st.session_state.praesenzarbeitsort_wert = st.session_state.praesenzarbeitsort_wert_day * st.session_state.anz_telearbeitstage
st.session_state.telearbeitsort_wert = st.session_state.telearbeitsort_wert_day * st.session_state.anz_telearbeitstage
st.session_state.shared_desk_wert = st.session_state.shared_desk_wert_day * st.session_state.anz_telearbeitstage
values = update_values()

######################################################################################################################################
# Arbeitswege
data_arbeitsweg = ws_Arbeitsweg.get_all_records()
df_arbeitsweg = pd.DataFrame(data_arbeitsweg)
transportmittel = df_arbeitsweg["Transportmittel"].unique().tolist()

with st.expander("Arbeitswege spezifizieren"):
    transport = st.radio("Wählen Sie das Transportmittel, mit dem Sie den Arbeitsweg primär zurücklegen", transportmittel)
    km_arbeitsweg = st.number_input("Geben Sie die Länge des Arbeitsweges in km an (Hin- und Rückfahrt)", min_value=0)

    if st.button("Arbeitsweg-Emissionen berechnen"):
        emissions_transportmittel = df_arbeitsweg[df_arbeitsweg["Transportmittel"] == transport]["kg CO2e je km"].values[0]
        emissions_arbeitswege = km_arbeitsweg * emissions_transportmittel
        st.session_state.arbeitsweg_wert_day = -emissions_arbeitswege / 5
        st.session_state.arbeitsweg_wert = st.session_state.arbeitsweg_wert_day * anz_telearbeitstage
        arbeitsweg_index = categories.index('Arbeitsweg') 
        values[arbeitsweg_index] = st.session_state.arbeitsweg_wert

        st.session_state.induzierte_wege_wert_day = st.session_state.arbeitsweg_wert_day * st.session_state.induzierte_wege_wert_anteil
        st.session_state.induzierte_wege_wert = st.session_state.induzierte_wege_wert_day * anz_telearbeitstage
        values[categories.index('induzierte Wege')] = st.session_state.induzierte_wege_wert

        st.write(f"Durch Ihren Arbeitsweg werden täglich etwa {emissions_arbeitswege:.3f} kg CO2e emittiert.")
        st.write(f"An jedem Telearbeitstag sparen sie dadurch Emissionen in Höhe von {st.session_state.arbeitsweg_wert_day:.3f} kg CO2e gegenüber der vollständigen Präsenzarbeit ein.")
        st.write(f"In Ihrem Telearbeitsmodell ergibt dies eine Reduktion des Treibhauspotenzials von {st.session_state.arbeitsweg_wert:.3f} kg CO2e gegenüber der vollständigen Präsenzarbeit im Beriech der Arbeitswege.")

######################################################################################################################################
# induzierte Wege

with st.expander("Anteil induzierte Wege verändern"):
    st.write(f"Typischerweise wird die 0,62-fache Länge des Arbeitsweges an Telearbeitstagen dennoch zurückgelegt. Hier können Sie diese Annahme verändern: ")
    st.session_state.induzierte_wege_wert_anteil = st.slider("Anteil induzierter Werte", min_value=0.005, max_value=1.19, value=0.001)
    if st.button("induzierte Wege neu berechnen"):
        st.session_state.induzierte_wege_wert_day = st.session_state.arbeitsweg_wert_day * -st.session_state.induzierte_wege_wert_anteil
        st.session_state.induzierte_wege_wert = st.session_state.induzierte_wege_wert_day * anz_telearbeitstage
        values[categories.index('induzierte Wege')] = st.session_state.induzierte_wege_wert
        st.write(f"Der Anteil induzierter Wege wurde neu berechnet.")

######################################################################################################################################
# Primärenergiebedarf Präsenzarbeitsort
data_ppg = ws_ppg.get_all_records() #ppg: Primärenergiebedarf Präsenzarbeitsort Gebäude
df_ppg = pd.DataFrame(data_ppg)
gebaeudekategorie = df_ppg["Gebäudekategorie"].unique().tolist()

data_pph = ws_pph.get_all_records() #pph: Primärenergiebedarf_Präsenzarbeitsort_Heizen
df_pph = pd.DataFrame(data_pph)
pph_kategorie = df_pph["Datensatz"].unique().tolist()

with st.expander("Primärenergiebedarf am Präsenzarbeitsort festlegen"):
    st.write(f"An Telearbeitstagen entfällt der Primärenergiebedarf am Präsenzarbeitsplatz. Hier können Sie diesen genauer spezifizieren")    

    flaeche_praesenzarbeitsplatz = st.number_input("Geben Sie die Fläche des Arbeitsplatzes in m2 an", min_value=9.0, value=13.75)

    gebaeudeauswahl = st.radio("Falls bekannt, Wählen Sie die Gebäudekategorie, welche ihren Präsenzarbeitsort am besten beschreibt", gebaeudekategorie)
    endenergiebedarf_Brennstoffe_ppg = df_ppg[df_ppg["Gebäudekategorie"] == gebaeudeauswahl]["Endenergiebedarf Brennstoffe"].values[0]
    endenergiebedarf_Strom_ppg = df_ppg[df_ppg["Gebäudekategorie"] == gebaeudeauswahl]["Endenergiebedarf Strom"].values[0]
    primaerenergiebedarf_Brennstoffe_ppg = df_ppg[df_ppg["Gebäudekategorie"] == gebaeudeauswahl]["Primärenergiebedarf"].values[0]
    if st.button("Primärenergiebedarf aus Gebäudekategorie berechnen"):
        st.session_state.praesenzarbeitsort_wert_day = -primaerenergiebedarf_Brennstoffe_ppg * flaeche_praesenzarbeitsplatz / 235 #235 Arbeitstage im Jahr
        st.session_state.praesenzarbeitsort_wert = st.session_state.praesenzarbeitsort_wert_day * anz_telearbeitstage
        values[categories.index('Primärenergiebedarf Präsenzarbeitsort')] = st.session_state.praesenzarbeitsort_wert
        st.write(f"Der Anteil induzierter Wege wurde anhand der Gebäudekategorie neu berechnet.")

    pph_auswahl = st.radio("Falls bekannt, Wählen Sie die Art, wie ihr Präsenzarbeitsort beheizt wird", pph_kategorie)
    if st.button("Primärenergiebedarf aus Gebäudekategorie und Heizdatensatz berechnen"):
        pph_wert = df_pph[df_pph["Datensatz"] == pph_auswahl]["CO2e je kWh"].values[0]
        st.session_state.praesenzarbeitsort_wert_day = -((endenergiebedarf_Brennstoffe_ppg * pph_wert + endenergiebedarf_Strom_ppg * 0.705) * flaeche_praesenzarbeitsplatz / 235 ) #0.705 Standardwert Strommix; 235 Arbeitstage Gesamtjahresbeizung im Jahr als Annahme
        st.session_state.praesenzarbeitsort_wert = st.session_state.praesenzarbeitsort_wert_day * anz_telearbeitstage
        values[categories.index('Primärenergiebedarf Präsenzarbeitsort')] = st.session_state.praesenzarbeitsort_wert
        st.write(f"Der Anteil induzierter Wege wurde anhand der Gebäudekategorie und Heizdatensatz neu berechnet.")



######################################################################################################################################
# Primärenergiebedarf Telearbeitsort 

data_pth = ws_pth.get_all_records() #pph: Primärenergiebedarf_Telearbeitsort_Heizen
df_pth = pd.DataFrame(data_pth)
pth_kategorie = df_pth["Datensatz"].unique().tolist()

with st.expander("Primärenergiebedarf am Telearbeitsort festlegen"):
    st.write(f"Wenn Sie die Temperatur während der Abwesenheit am Telearbeitsplatz durch Präsenzarbeit absenken, so ergeben sich mehr Emissionen an Telearbeitstagen ")
    temperaturabsenkung = st.radio("Senken Sie die Temperatur am Telearbeitsplatz während der Präsenzarbeitszeiten?", ("Ja", "Nein"))
    if temperaturabsenkung == "Nein":
        st.session_state.telearbeitsort_wert_day = 0
        st.session_state.telearbeitsort_wert = st.session_state.telearbeitsort_wert_day * anz_telearbeitstage
        values[categories.index('Primärenergiebedarf Telearbeitsort')] = st.session_state.telearbeitsort_wert
    else:
        st.write(f"Hinweis: Die Angaben zu Endenergiebedarf und Gebäudenutzfläche finden Sie im Energieausweis.")
        endenergiebedarf_Brennstoffe_pt = st.number_input("Wie hoch ist der Endenergiebedarf ihrer Wohnung in kWh/(m²a)? ", min_value=0, value=115) 
        flaeche_telearbeitsplatz = st.number_input("Wie hoch ist die Gebäudenutzfläche im m²? (alternativ: Wie hoch ist die Fläche, die während der Abwesenheit weniger beheizt wird?)", min_value=0, value=65)
        pth_auswahl = st.radio("Wie wird geheizt?", pth_kategorie)
        pth_wert = df_pth[df_pth["Datensatz"] == pth_auswahl]["kg CO2e je kWh"].values[0]
    
    if st.button("Primärenergiebedarf am Telearbeitsort neu berechnen"):
        if temperaturabsenkung == "Ja":
            st.session_state.telearbeitsort_wert_day = pth_wert * flaeche_telearbeitsplatz * endenergiebedarf_Brennstoffe_pt / 365 / 17 * 9.5 * 0.16 #365 Tage im Jahr heizen bei Wohngebäuden, an 17 Stunden/d, mit 9,5 Abwesenheitsstunden, 0.16 Faktor der Temperaturabsenkung um 4°K
        st.session_state.telearbeitsort_wert = st.session_state.telearbeitsort_wert_day * anz_telearbeitstage
        values[categories.index('Primärenergiebedarf Telearbeitsort')] = st.session_state.telearbeitsort_wert
    


######################################################################################################################################
# end charts / Gesamtauswertung
######################################################################################################################################
st.title("Ergebnisse")

# bar chart: categories preparations
df_chart = pd.DataFrame({
    'Kategorie': categories, 
    'Wert': values 
})
df_chart['Farbe'] = df_chart['Wert'].apply(lambda x: 'green' if x < 0 else 'red') #positive Werte rot, negative grün
fig_categories = px.bar( 
    df_chart, 
    x = 'Kategorie',
    y='Wert', 
    title="Kategorien: Veränderungen auf das Treibhauspotenzial je Telearbeitstag [kg CO2e/d]", 
    color='Farbe', 
    color_discrete_map={'green': 'green', 'red': 'red'}, 
    text='Wert' # Datenbeschriftungen hinzufügen 
)
fig_categories.update_layout( 
    xaxis_title=None, 
    yaxis_title=None, 
    xaxis_tickangle=-45, 
    showlegend=True, 
    legend_title_text='Veränderungen:', 
    legend=dict(itemsizing='constant', title_font_size=16, font_size=14) 
)
fig_categories.update_traces(texttemplate='%{text:.2f}', textposition='outside')
fig_categories.update_layout(showlegend=True, legend_title_text='Veränderungen:', legend=dict(itemsizing='constant', title_font_size=16, font_size=14))
newnames = {'green': 'Emissionsreduktion durch Telearbeit', 'red': 'Emissionserhöhung durch Telearbeit'} 
fig_categories.for_each_trace(lambda t: t.update(name = newnames[t.name]))

# bar chart sum values (Summe der Kategorien) preparations
sum_values = sum(values)
color_sum = 'green' if sum_values < 0 else 'red'

fig_sum = px.bar(
    x=[sum_values], 
    y=[''], 
    text=[sum_values], 
    title=None, 
    orientation='h',
    color_discrete_sequence=[color_sum] 
)
max_value = max(abs(sum_values), 1)
fig_sum.update_layout( xaxis_title="[CO2e je Arbeitstag gegenüber vollständiger Präsenzarbeit]", yaxis_title=None, height=200, showlegend=False, yaxis=dict(showticklabels=False),
    xaxis=dict( range=[-max_value*2, max_value*2], zeroline=True, zerolinewidth=2, zerolinecolor='black', gridcolor='LightGray' ))

left_text = "<b>Telearbeit ist zu bevorzugen</b>" if sum_values < 0 else "Telearbeit ist zu bevorzugen" 
right_text = "<b>Präsenzarbeit ist zu bevorzugen</b>" if sum_values > 0 else "Präsenzarbeit ist zu bevorzugen"
fig_sum.add_annotation( x=0.25, y=0.5, text=left_text, showarrow=False, xref="paper", yref="paper", align="right", xanchor='right' ) 
fig_sum.add_annotation( x=0.75, y=0.5, text=right_text, showarrow=False, xref="paper", yref="paper", align="left", xanchor='left' )

# plot charts
st.plotly_chart(fig_sum)
with st.expander("Aufschlüsselung nach Kategorien"):
    st.plotly_chart(fig_categories)