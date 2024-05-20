import pandas as pd
import numpy as np
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import plotly.express as px
import calendar
import os
import glob
import re


# ------------------RELEVANT Auswertung Erfüllungsgrad nach Präferenzen (aggregiert über alle MA)-----------------


# Einlesen der Excel-Datei
df = pd.read_excel('merged_dataframe.xlsx')

# Entfernen der Spalte 'Datum', falls vorhanden
if 'Datum' in df.columns:
    df.drop('Datum', axis=1, inplace=True)

# Ersetzen von WAHR/FALSCH und anderen nicht zählbaren Einträgen durch numerische Werte für die Visualisierung
df.replace({
    'WAHR': 1,
    'FALSCH': 0,
    'keine beliebten Dienste eingetragen': np.nan,
    'keine unbeliebten Dienste eingetragen': np.nan,
    '': np.nan  # Leere Strings zu NaN konvertieren, falls vorhanden
}, inplace=True)

# Schmelzen des DataFrames, um ein "long format" zu erstellen
df_melted = df.melt(id_vars=['MA-ID'], var_name='Kriterien', value_name='Erfüllt')

# Filtern der Zeilen, die die spezifischen Werte in der Spalte 'Erfüllt' enthalten
df_filtered = df_melted[~df_melted['Erfüllt'].isin(
    ['keine beliebten Dienste eingetragen', 'keine unbeliebten Dienste eingetragen', np.nan])]

# Ersetzen der Werte durch Booleans für einfache Berechnung
df_filtered.loc[:, 'Erfüllt'] = df_filtered['Erfüllt'].astype(bool)

# Gruppierung nach Kriterien und Berechnung des Prozentsatzes der wahren Werte über alle MA-ID
df_grouped_criteria = df_filtered.groupby('Kriterien')['Erfüllt'].mean() * 100
df_grouped_criteria = df_grouped_criteria.reset_index()

# Einrichten des Balkendiagramms
fig_criteria = go.Figure()

# Hinzufügen der Daten zum Balkendiagramm
fig_criteria.add_trace(go.Bar(
    x=df_grouped_criteria['Kriterien'],
    y=df_grouped_criteria['Erfüllt'],
    marker_color='blue'  # Farbe der Balken anpassen
))

# Layout anpassen
fig_criteria.update_layout(
    title='Erfüllungsgrad nach Präferenzen (aggregiert über alle MA und gesamter Zeitraum)',
    xaxis_title='Präferenzen',
    yaxis_title='Erfüllungsgrad (%)'
)

# Anzeigen des Plots
fig_criteria.show()

 #Speichern als HTML-Datei
fig_criteria.write_html("Erfüllungsgrad nach Präferenzen (aggregiert über alle MA und gesamter Zeitraum).html")


# -----------------------RELEVANT Erfüllungsgrad der Präferenzen über die Zeit (aggregiert über alle MA)------------------------------


# Einlesen der Excel-Datei
df = pd.read_excel('merged_dataframe.xlsx')

# Entfernen der Spalte 'Datum', falls vorhanden
if 'Datum' in df.columns:
    # Schmelzen des DataFrames, um ein "long format" zu erstellen
    df_melted = df.melt(id_vars=['MA-ID', 'Datum'], var_name='Kriterien', value_name='Erfüllt')

    # Filtern der Zeilen, die die spezifischen Werte in der Spalte 'Erfüllt' enthalten
    df_filtered = df_melted[~df_melted['Erfüllt'].isin(
        ['keine beliebten Dienste eingetragen', 'keine unbeliebten Dienste eingetragen', np.nan])]

    # Ersetzen der Werte durch Booleans für einfache Berechnung
    df_filtered['Erfüllt'] = df_filtered['Erfüllt'].astype(bool)

    # Monat und Jahr aus dem Datum extrahieren
    df_filtered['Monat'] = df_filtered['Datum'].dt.strftime('%B %Y')

    # Sortieren der Monate in der richtigen Reihenfolge
    months_order = pd.date_range(start=df_filtered['Datum'].min(), end=df_filtered['Datum'].max(), freq='M').strftime('%B %Y')
    df_filtered['Monat'] = pd.Categorical(df_filtered['Monat'], categories=months_order, ordered=True)

    # Gruppierung nach Kriterien und Monat und Berechnung des Prozentsatzes der wahren Werte
    df_grouped_criteria_month = df_filtered.groupby(['Kriterien', 'Monat'])['Erfüllt'].mean() * 100
    df_grouped_criteria_month = df_grouped_criteria_month.reset_index()

    # Einrichten der Subplots
    fig = make_subplots(rows=len(df_grouped_criteria_month['Kriterien'].unique()), cols=1,
                        subplot_titles=df_grouped_criteria_month['Kriterien'].unique())

    # Hinzufügen der Daten zu den Subplots
    for i, (kriterium, data) in enumerate(df_grouped_criteria_month.groupby('Kriterien')):
        fig.add_trace(
            go.Scatter(x=data['Monat'], y=data['Erfüllt'], mode='markers+lines', name=kriterium),
            row=i + 1, col=1
        )

    # Layout anpassen
    fig.update_layout(
        title='Durchschnittlicher Erfüllungsgrad der Präferenzen über die Zeit (aggregiert über alle MA)',
        xaxis_title='Monat',
        yaxis_title='Erfüllungsgrad (%)',
        showlegend=True,
        height=600 * len(df_grouped_criteria_month['Kriterien'].unique())  # Höhe des Diagramms basierend auf der Anzahl der Kriterien
    )

    # Beschriftung aller Y-Achsen
    for i in range(1, len(df_grouped_criteria_month['Kriterien'].unique()) + 1):
        fig.update_yaxes(title_text='Erfüllungsgrad (%)', row=i, col=1)

    # Anzeigen des Plots
    fig.show()
else:
    print("Die Spalte 'Datum' ist nicht im DataFrame vorhanden.")
#
# # # Speichern als HTML-Datei
# fig.write_html("Durchschnittlicher Erfüllungsgrad der Präferenzen über die Zeit (aggregiert über alle MA).html")



#__________________________________RELEVANT Tabelle pro Monat mit Präferenzerfüllung und Abweichung pro MA-ID________________________
def prepare_dataframe(file_path):
    # Einlesen der Excel-Datei
    df = pd.read_excel(file_path)

    # Vorbereitung der Daten wie im obigen Code
    if 'Datum' in df.columns:
        df.drop('Datum', axis=1, inplace=True)

    df.replace({
        'WAHR': 1,
        'FALSCH': 0,
        'keine beliebten Dienste eingetragen': np.nan,
        'keine unbeliebten Dienste eingetragen': np.nan,
        '': np.nan
    }, inplace=True)

    # Löschen der spezifischen Spalten
    columns_to_drop = [
        'Anzahl Dienste pro Block',
        'Anzahl Dienste pro Block Abweichung vom Durchschnitt',
        'Opt. Anz. Dienste in Block',
        'Opt. Anz. Dienste in Block Abweichung vom Durchschnitt'
    ]
    df.drop(columns=columns_to_drop, errors='ignore', inplace=True)

    df_melted = df.melt(id_vars=['MA-ID'], var_name='Kriterien',
                        value_name='Erfüllt')  # 'MA-ID' als id_vars hinzugefügt
    df_filtered = df_melted[~df_melted['Erfüllt'].isna()]
    df_filtered = df_filtered.copy()  # Erstelle eine Kopie, um SettingWithCopyWarning zu vermeiden
    df_filtered['Erfüllt'] = df_filtered['Erfüllt'].astype(bool)

    df_grouped = df_filtered.groupby(['MA-ID', 'Kriterien'])['Erfüllt'].mean() * 100
    average_per_criteria = df_grouped.groupby('Kriterien').mean()
    df_pivot = df_grouped.reset_index().pivot(index='MA-ID', columns='Kriterien', values='Erfüllt')

    df_pivot.loc['Durchschnitt'] = average_per_criteria

    df_pivot_with_deviation = df_pivot.copy()
    for col in df_pivot.columns:
        deviation_col = col + ' Abweichung vom Durchschnitt'
        df_pivot_with_deviation[deviation_col] = df_pivot[col] - df_pivot.loc['Durchschnitt', col]

    ordered_columns = []
    for col in df_pivot.columns:
        ordered_columns.append(col)
        ordered_columns.append(col + ' Abweichung vom Durchschnitt')

    df_pivot_with_deviation = df_pivot_with_deviation[ordered_columns]

    return df_pivot_with_deviation

breakpoint()

# Verzeichnis mit den Excel-Dateien
directory = 'Dataframes'

# Liste der Dateinamen im Verzeichnis, die mit "df_" beginnen
file_names = [file for file in os.listdir(directory) if file.startswith('df_')]

# DataFrame für jeden Monat erstellen und in einem Dictionary speichern
dfs = {}
for file_name in file_names:
    file_path = os.path.join(directory, file_name)
    month_df = prepare_dataframe(file_path)
    dfs[file_name] = month_df



# # Beispiel: Zugriff auf das DataFrame für einen bestimmten Monat
# januar_df = dfs['df_Januar 2023.xlsx']

# Beispiel: Speichern der DataFrames in Excel-Dateien
output_directory = 'Dataframes/DFs Tabelle'

for file_name, df in dfs.items():
    output_file_name = f'Tabelle_{file_name}'  # Beispiel für den Dateinamen
    output_file_path = os.path.join(output_directory, output_file_name)
    df.to_excel(output_file_path, index=False)
    df.reset_index().to_excel(output_file_path, index=False)




# _________________________RELEVANT Mehrer Plots die, die Abweichung zum Durchschnitt anzeigen als Subplot, der sich nur auf ein Monat bezieht------------

# Pfad zu dem Verzeichnis, in dem die Dateien gespeichert sind
path = 'Dataframes/DFs Tabelle/*.xlsx'  # Anpassen an Ihren spezifischen Pfad

# Liste aller Excel-Dateien, die mit 'df_' beginnen
file_list = glob.glob(path)

# Liste der Spalten für die Visualisierung
columns = [
    'Arbeitsblocklänge Opt high und Opt low erfüllt',
    'Beachtung unbeliebte Dienste',
    'Erfüllung beliebte Dienste',
    'Opt aufeinanderf. WE mit Arbeit beachtet',
    'Opt. Anz. Dienste in Block erfüllt',
    'Opt. freie Weekends erfüllt'
]

# Durchlaufen jeder Datei in der Liste
for file in file_list:
    # Daten laden
    data = pd.read_excel(file)

    # Entfernen der Zeile "Durchschnitt", falls vorhanden
    data = data[data['MA-ID'] != 'Durchschnitt']
    data['MA-ID'] = data['MA-ID'].astype(str)  # MA-ID als String

    # Extrahieren von Monat und Jahr aus dem Dateinamen
    match = re.search(r'df_(\w+) (\d{4})', file)
    if match:
        month_year = match.group(1) + " " + match.group(2)
    else:
        month_year = "Unbekanntes Datum"

    # Erstellen der Subplots
    fig = make_subplots(rows=len(columns), cols=1, subplot_titles=columns, vertical_spacing=0.05)

    for i, column in enumerate(columns):
        row = i + 1
        col = 1
        average_value = data[column].mean()
        fig.add_trace(go.Scatter(x=data['MA-ID'], y=data[column], mode='markers', name=column), row=row, col=col)
        fig.add_hline(y=average_value, line_dash='dash', line_color='red', annotation_text="Durchschnitt", row=row, col=col)

    fig.update_layout(height=1800, width=1500, title_text=f"Übersicht der Erfüllungskriterien pro MA-ID für {month_year}")
    for i in range(1, len(columns) + 1):
        fig.update_xaxes(title_text="MA-ID", row=i, col=1)
        fig.update_yaxes(title_text="Erfüllungsgrad (%)", row=i, col=1)

    # Speichern oder Anzeigen des Plots
    # fig.write_html(f'{file}_plot_Abweichung vom Durchschnitt.html')  # Speichern als HTML-Datei
    fig.show()




# ----------------------------RELEVANT Tabellen mit Erfüllungsgrad_________________________

# Pfad zu dem Verzeichnis, in dem die Dateien gespeichert sind
path = 'Dataframes/DFs Tabelle/*.xlsx'

# Liste aller Excel-Dateien im Verzeichnis
file_list = glob.glob(path)

# Durchlaufen jeder Datei in der Liste
for file in file_list:
    df = pd.read_excel(file)

    # Monat und Jahr aus dem Dateinamen extrahieren
    match = re.search(r'Tabelle_df_(\w+) (\d{4})\.xlsx', file)
    if match:
        month_year = match.group(1) + " " + match.group(2)
    else:
        month_year = "Unbekanntes Datum"

    # Farben für Kopfzeilen definieren
    header_colors = {
        'MA-ID': 'rgb(95,158,160)',
        'Arbeitsblocklänge Opt high und Opt low erfüllt': 'rgb(0, 255, 0)',
        'Beachtung unbeliebte Dienste': 'rgb(0, 255, 0)',
        'Erfüllung beliebte Dienste': 'rgb(0, 255, 0)',
        'Opt aufeinanderf. WE mit Arbeit beachtet': 'rgb(0, 255, 0)',
        'Opt. Anz. Dienste in Block erfüllt': 'rgb(0, 255, 0)',
        'Opt. freie Weekends erfüllt': 'rgb(0, 255, 0)',
        'Arbeitsblocklänge Opt high und Opt low erfüllt Abweichung vom Durchschnitt': 'rgb(30,144,255)',
        'Beachtung unbeliebte Dienste Abweichung vom Durchschnitt': 'rgb(30,144,255)',
        'Erfüllung beliebte Dienste Abweichung vom Durchschnitt': 'rgb(30,144,255)',
        'Opt aufeinanderf. WE mit Arbeit beachtet Abweichung vom Durchschnitt': 'rgb(30,144,255)',
        'Opt. Anz. Dienste in Block erfüllt Abweichung vom Durchschnitt': 'rgb(30,144,255)',
        'Opt. freie Weekends erfüllt Abweichung vom Durchschnitt': 'rgb(30,144,255)',
    }

    # Erstellen der Plotly-Tabelle für jede Datei
    fig = go.Figure(data=[go.Table(
        header=dict(values=df.columns,
                    fill_color=[header_colors.get(col, 'lavender') for col in df.columns],
                    # Default color if not specified
                    align='left'),
        cells=dict(values=[df[col] for col in df.columns],
                   fill_color='lavender',
                   align='left'))
    ])

    # Titel der Tabelle anpassen
    fig.update_layout(
        title=f'Erfüllungsgrad pro Mitarbeiter:in und Präferenz mit Abweichung vom Durchschnitt für {month_year}',
        width=2500,
        height=1000
    )

    # Speichern der Tabelle als HTML-Datei
    output_filename = f'Tabelle_Erfüllungsgrad_{month_year}.html'  # Dateiname anpassen
    output_filepath = os.path.join('Dataframes/Plots', output_filename)  # Output-Verzeichnispfad anpassen
    fig.write_html(output_filepath)

# _____________________________RELEVANT Plot Durchschnittlicher Erfüllungsgrad pro MA nach Präferenzen über den gesamten Zeitraum__________

# Einlesen der Excel-Datei
df = pd.read_excel('merged_dataframe.xlsx')

# Entfernen der Spalte 'Datum', falls vorhanden
if 'Datum' in df.columns:
    df.drop('Datum', axis=1, inplace=True)

# Ersetzen von WAHR/FALSCH und anderen nicht zählbaren Einträgen durch numerische Werte für die Visualisierung
df.replace({
    'WAHR': 1,
    'FALSCH': 0,
    'keine beliebten Dienste eingetragen': np.nan,
    'keine unbeliebten Dienste eingetragen': np.nan,
    '': np.nan  # Leere Strings zu NaN konvertieren, falls vorhanden
}, inplace=True)

# Schmelzen des DataFrames, um ein "long format" zu erstellen
df_melted = df.melt(id_vars=['MA-ID'], var_name='Kriterien', value_name='Erfüllt')

# Filtern der Zeilen, die die spezifischen Werte in der Spalte 'Erfüllt' enthalten
df_filtered = df_melted[~df_melted['Erfüllt'].isin(
    ['keine beliebten Dienste eingetragen', 'keine unbeliebten Dienste eingetragen', np.nan])]

# Ersetzen der Werte durch Booleans für einfache Berechnung
df_filtered.loc[:, 'Erfüllt'] = df_filtered['Erfüllt'].astype(bool)

# Gruppierung nach Kriterien und Berechnung des Prozentsatzes der wahren Werte über alle MA-ID
df_grouped = df_filtered.groupby(['Kriterien', 'MA-ID'])['Erfüllt'].mean() * 100
df_grouped = df_grouped.reset_index()

# Einrichten der Subplots
fig = make_subplots(rows=len(df_grouped['Kriterien'].unique()), cols=1,
                    subplot_titles=df_grouped['Kriterien'].unique(), vertical_spacing=0.05)

# Alle MA-IDs sammeln, um sicherzustellen, dass alle MA-IDs auf der X-Achse erscheinen
all_ma_ids = df_grouped['MA-ID'].unique()

# Iteration über jedes Kriterium für die Subplots
for i, kriterium in enumerate(df_grouped['Kriterien'].unique(), start=1):
    # Daten für das aktuelle Kriterium auswählen
    kriterium_data = df_grouped[df_grouped['Kriterien'] == kriterium]

    # Hinzufügen der Balken für das aktuelle Kriterium
    fig.add_trace(
        go.Bar(x=kriterium_data['MA-ID'], y=kriterium_data['Erfüllt'], name=kriterium),
        row=i, col=1
    )

    # Beschriftung der X-Achse für jeden Subplot explizit setzen
    fig.update_xaxes(
        title_text="MA-ID",  # Titel der X-Achse für jeden Subplot
        tickmode='array',
        tickvals=all_ma_ids,  # Stelle sicher, dass alle MA-IDs angezeigt werden
        ticktext=[str(id) for id in all_ma_ids],  # Konvertiere IDs zu Strings für die Anzeige
        row=i, col=1
    )

    # Beschriftung der Y-Achse
    fig.update_yaxes(
        title_text="Erfüllungsgrad (%)",  # Titel der Y-Achse
        row=i, col=1
    )

# Layout anpassen
fig.update_layout(
    title='Durchschnittlicher Erfüllungsgrad pro MA nach Präferenzen über den gesamten Zeitraum',
    showlegend=True,
    font=dict(size=12),
    height=300 * len(df_grouped['Kriterien'].unique())  # Höhe anpassen, abhängig von der Anzahl der Subplots
)

# Anzeigen des Plots
fig.show()

# fig.write_html('Plot_Durchschnittlicher Erfüllungsgrad pro MA nach Präferenzen über den gesamten Zeitraum.html')

# ____________________________________RELEVANT Tabelle Durchschnittlicher Erfüllungsgrad pro MA nach Präferenzen über den gesamten Zeitraum___________

# Lese die Daten ein
df = pd.read_excel('merged_dataframe.xlsx')

# Entferne die Spalte 'Datum', falls vorhanden
if 'Datum' in df.columns:
    df.drop('Datum', axis=1, inplace=True)

# Ersetze Werte durch numerische Werte für die Visualisierung
df.replace({
    'WAHR': 1,
    'FALSCH': 0,
    'keine beliebten Dienste eingetragen': np.nan,
    'keine unbeliebten Dienste eingetragen': np.nan,
    '': np.nan
}, inplace=True)

# Schmelze den DataFrame, um ein "long format" zu erstellen
df_melted = df.melt(id_vars=['MA-ID'], var_name='Kriterien', value_name='Erfüllt')

# Filtere die Zeilen, die spezifische Werte in der Spalte 'Erfüllt' enthalten
df_filtered = df_melted[~df_melted['Erfüllt'].isin(
    ['keine beliebten Dienste eingetragen', 'keine unbeliebten Dienste eingetragen', np.nan])]

# Ersetze die Werte durch Booleans für einfache Berechnung
df_filtered['Erfüllt'] = df_filtered['Erfüllt'].astype(bool)

# Gruppiere nach Kriterien und berechne den Prozentsatz der wahren Werte über alle MA-ID
df_grouped = df_filtered.groupby(['Kriterien', 'MA-ID'])['Erfüllt'].mean() * 100
df_grouped = df_grouped.reset_index()

# Einrichte die Plotly-Tabelle
fig = go.Figure()

# Alle einzigartigen Kriterien ermitteln
unique_kriterien = df_grouped['Kriterien'].unique()
num_kriterien = len(unique_kriterien)

# Definiere eine Funktion, um die Farbe für jede Zelle basierend auf dem Wert zu bestimmen
def get_cell_color(value, lowest_values):
    if value in lowest_values:
        return 'lightcoral'  # Farbe für die niedrigsten Werte
    else:
        return 'lavender'  # Standardfarbe

# Für jedes Kriterium eine Tabelle erstellen
for i, kriterium in enumerate(unique_kriterien):
    kriterium_data = df_grouped[df_grouped['Kriterien'] == kriterium]

    # Sortiere die Daten nach MA-ID, um eine konsistente Darstellung zu gewährleisten
    kriterium_data = kriterium_data.sort_values('MA-ID')

    # Finde die vier niedrigsten Werte
    lowest_values = kriterium_data.nsmallest(4, 'Erfüllt')['Erfüllt'].tolist()

    # Berechne den Durchschnittswert für diese Tabelle
    average_value = kriterium_data['Erfüllt'].mean()

    # Füge die Durchschnittszeile hinzu
    kriterium_data.loc[len(kriterium_data.index)] = ['Durchschnitt', 'Durchschnitt', average_value]

    # Tabelle erstellen, Titel für jede Tabelle hinzufügen
    fig.add_trace(
        go.Table(
            header=dict(values=[f'MA-ID - {kriterium}', f'Erfüllungsgrad (%) - {kriterium}'],
                        fill_color='paleturquoise',
                        align='left'),
            cells=dict(values=[kriterium_data['MA-ID'],
                                kriterium_data['Erfüllt'].round(2)],
                       fill_color=[[get_cell_color(value, lowest_values) for value in kriterium_data['Erfüllt']]],
                       align='left'),
            domain=dict(x=[0, 1], y=[1 - (i + 1) / num_kriterien, 1 - i / num_kriterien])  # Anpassen der Domain für jede Tabelle
        )
    )

# Layout anpassen
fig.update_layout(
    title='Durchschnittlicher Erfüllungsgrad pro MA nach Präferenzen über den gesamten Zeitraum',
    showlegend=False,
    height=100 * 5 * (num_kriterien + 1)  # Höhe anpassen, abhängig von der Anzahl der Kriterien und der Durchschnittszeile
)

# Anzeigen der Tabelle
fig.show()

# fig.write_html('Tabelle_Durchschnittlicher Erfüllungsgrad pro MA nach Präferenzen über den gesamten Zeitraum.html')