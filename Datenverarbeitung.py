import pandas as pd
import json
import openpyxl
import plotly.express as px
import numpy as np
from datetime import datetime
import calendar



# Datei laden

with open("Januar 2024/Januar 2024.json") as file:
    data = json.load(file)


#____________________________Monatsbeginn und Ende identifizieren_________________________________

# Monatsnamen in Zahlen umwandeln
def month_to_number(month_name):
    months = {
        "Januar": 1, "Februar": 2, "März": 3, "April": 4, "Mai": 5, "Juni": 6,
        "Juli": 7, "August": 8, "September": 9, "Oktober": 10, "November": 11, "Dezember": 12
    }
    return months.get(month_name)

def extract_monthly_data(file_path):
    # Jahr und Monat aus dem Dateipfad extrahieren
    parts = file_path.split('/')
    month_name, year = parts[-1].split()  # Das letzte Element des Pfades aufsplitten
    month = month_to_number(month_name)
    year = int(year.split('.')[0])  # '.json' aus dem Jahr entfernen

    # Den letzten Tag des Monats ermitteln
    last_day = calendar.monthrange(year, month)[1]

    # Datumsgrenzen festlegen
    start_date = datetime(year, month, 1)
    end_date = datetime(year, month, last_day)

    # Datei laden
    with open(file_path) as file:
        data = json.load(file)

    # Daten von JSON-File extrahieren
    extracted_data = []
    for item in data['shiftAssignmentsByEmploymentIdAndShiftId']:
        ma_id = item['first']
        date_service = item['second'].replace('_', '').split('-')  # Unterstriche entfernen vor dem Aufteilen
        date_str = '-'.join(date_service[:3])
        date = datetime.strptime(date_str, '%Y-%m-%d')
        service = '-'.join(date_service[3:])

        # Überprüfen, ob das Datum innerhalb des festgelegten Zeitraums liegt
        if start_date <= date <= end_date:
            extracted_data.append([ma_id, date_str, service])

    # DataFrame erstellen
    return pd.DataFrame(extracted_data, columns=['MA-ID', 'Datum', 'Dienst'])


# Anwendung des geänderten Codes
df = extract_monthly_data("Januar 2024/Januar 2024.json")

#____________________________Json anpassen_________________________________

# Wochentag basierend auf dem Datum hinzufügen
df['Wochentag'] = pd.to_datetime(df['Datum']).dt.strftime('%A')

# Neue Spalte "Funktion" hinzufügen
df['Funktion'] = df['Dienst'].str.split('-').str[-1]

# "-L" aus der Spalte "Dienst" entfernen
df['Dienst'] = df['Dienst'].str.replace('-L', '')

# "-Dipl" aus der Spalte "Dienst" entfernen
df['Dienst'] = df['Dienst'].str.replace('-Dipl', '')

# "-FaGe" aus der Spalte "Dienst" entfernen
df['Dienst'] = df['Dienst'].str.replace('-FaGe', '')

# "-MA" aus der Spalte "Dienst" entfernen
df['Dienst'] = df['Dienst'].str.replace('-MA', '')

# Spaltenreihenfolge ändern
df = df[['MA-ID', 'Funktion', 'Datum', 'Wochentag', 'Dienst']]

# Konvertiererung Datum in Datums-Zeitstempel
df['Datum'] = pd.to_datetime(df['Datum'])

#____________________________Arbeitsblöcke identifizieren_________________________________

# Berechne die Differenz zwischen aufeinanderfolgenden Daten
df['Diff'] = df.groupby('MA-ID')['Datum'].diff()

# Ersetze negative Differenzen durch 0
df['Diff'] = df['Diff'].fillna(0)

# Setze 'Arbeitsblocklänge' auf 1, wenn 'Diff' gleich 1 Tag ist
df.loc[df['Diff'] == pd.Timedelta(days=1), 'Arbeitsblocklänge'] = 1

# Fülle die Spalte "Arbeitsblocklänge" mit dem entsprechenden Dienst
df['Arbeitsblockdienst'] = df['Dienst'].where(df['Arbeitsblocklänge'] == 1, '')

# Excel-Datei einlesen
dateipfad = 'Januar 2024/IPSP-Inputs_LAK_2OG_Januar2024.xlsm'
df_excel = pd.read_excel(dateipfad, sheet_name='ma-profil')

# Umbenennen der Spalten in df_excel
df_excel.rename(columns={'Unnamed: 0': 'ID',
                         'Dienstpräferenzen': 'Beliebte Dienste',
                         'Unnamed: 7': 'Unbeliebte Dienste',
                         'Arbeits- und Freiblock': 'Max. + Opt high + Opt-low Arbeitsblocklänge',
                         'Unnamed: 10': 'Min. + Opt. Freiblockdauer (in Std.)',
                         'Unnamed: 11': 'Max. + Opt. Anz. Dienste in Block',
                         'Unnamed: 13': 'Bestrafte Abfolgen',
                         'Wochenendarbeit': 'Min + Opt. freie Weekends',
                         'Unnamed: 20': 'Max + Opt aufeinanderf. WE mit Arbeit',
                         'Monatliche Dienstlimiten und Präferenzen': 'Max. + Opt. Anz. Dienste '},
                inplace=True)

# über die Spalten itterieren und  NaN-Werte mit den Werten der dritten Zeile füllen
for col in ['Max. + Opt high + Opt-low Arbeitsblocklänge',
            'Min. + Opt. Freiblockdauer (in Std.)',
            'Max. + Opt. Anz. Dienste in Block',
            'Min + Opt. freie Weekends',
            'Max + Opt aufeinanderf. WE mit Arbeit']:
    df_excel[col].fillna(df_excel[col][2], inplace=True)

# DataFrame für den Abgleich mit Excel-Daten erstellen
df_abgleich = df.copy()

# Spalte "Beliebte Dienste" und "Unbeliebte Dienste" aus Excel extrahieren und mit df_abgleich abgleichen
df_excel_beliebt_unbeliebt = df_excel[['ID', 'Beliebte Dienste', 'Unbeliebte Dienste']].copy()
df_excel_beliebt_unbeliebt.rename(columns={'ID': 'MA-ID'}, inplace=True)
df_abgleich = pd.merge(df_abgleich, df_excel_beliebt_unbeliebt, on='MA-ID', how='left')


# Funktion zum Extrahieren von Diensttypen und Übersetzen der Wochentage
def extrahiere_diensttyp_und_wochentage(dienste):
    diensttypen = []
    wochentage = []
    wochentag_übersetzung = {
        'Mo': 'Monday',
        'Di': 'Tuesday',
        'Mi': 'Wednesday',
        'Do': 'Thursday',
        'Fr': 'Friday',
        'Sa': 'Saturday',
        'So': 'Sunday'
    }

    # Dienste aufteilen und analysieren
    if pd.isna(dienste):
        return '', ''
    for eintrag in dienste.split(','):
        eintrag = eintrag.strip()
        if '-' in eintrag:
            dienst, tage = eintrag.split('-')
            diensttypen.append(dienst.strip())
            wochentage.append(' '.join([wochentag_übersetzung.get(tag.strip(), '') for tag in tage.split()]))
        else:
            diensttypen.append(eintrag)
            wochentage.append('')

    return ', '.join(diensttypen), '; '.join(wochentage)


# Anpassung der Spalten "Beliebte Dienste" und "Unbeliebte Dienste"
for index, row in df_abgleich.iterrows():
    diensttypen, wochentage = extrahiere_diensttyp_und_wochentage(row['Beliebte Dienste'])
    df_abgleich.at[index, 'Beliebte Dienste'] = diensttypen
    df_abgleich.at[index, 'Wochentage beliebte Dienste'] = wochentage

    diensttypen, wochentage = extrahiere_diensttyp_und_wochentage(row['Unbeliebte Dienste'])
    df_abgleich.at[index, 'Unbeliebte Dienste'] = diensttypen
    df_abgleich.at[index, 'Wochentage unbeliebte Dienste'] = wochentage

# Spaltenreihenfolge anpassen
spalten = ['MA-ID', 'Funktion', 'Datum', 'Wochentag', 'Dienst', 'Beliebte Dienste', 'Wochentage beliebte Dienste',
           'Unbeliebte Dienste', 'Wochentage unbeliebte Dienste']
df_abgleich = df_abgleich[spalten]

#____________________________Überprüfung beliebter Dienste_________________________________

# Excel-Datei Lookuptable einlesen
dateipfad = 'Diensttypgruppen.xlsx'
df_diensttypgruppen = pd.read_excel(dateipfad)


# Initialisierung eines leeren Mappings
neues_dienst_mapping = {}

# Iteration über die Spaltennamen und Inhalte des DataFrame
for spaltenname in df_diensttypgruppen.columns:
    # Jede Kategorie (Spaltenname) wird jedem Diensttyp in dieser Kategorie zugeordnet
    for dienst in df_diensttypgruppen[spaltenname].dropna().unique():
        neues_dienst_mapping[dienst] = spaltenname


def pruefe_erfuellung_beliebter_dienste(row, dienst_mapping):
    beliebte_dienste = row['Beliebte Dienste']
    wochentage_beliebte_dienste = row['Wochentage beliebte Dienste']
    aktueller_dienst = row['Dienst']
    aktueller_wochentag = row['Wochentag']

    if pd.isna(beliebte_dienste) or beliebte_dienste.strip() == '':
        return "keine beliebten Dienste eingetragen"

    if aktueller_dienst in dienst_mapping:
        dienst_kategorie = dienst_mapping[aktueller_dienst]
        beliebte_dienste_kategorien = {dienst_mapping[dt]: wochentage for dt, wochentage in
                                       zip(beliebte_dienste.split(', '), wochentage_beliebte_dienste.split('; ')) if
                                       dt in dienst_mapping}

        if dienst_kategorie in beliebte_dienste_kategorien:
            wochentage = beliebte_dienste_kategorien[dienst_kategorie]
            if not wochentage.strip() or aktueller_wochentag in wochentage:
                return True
    return False

df_abgleich['Erfüllung beliebte Dienste'] = df_abgleich.apply(
    pruefe_erfuellung_beliebter_dienste, axis=1, dienst_mapping=neues_dienst_mapping
)




# Aktualisieren der Spaltenreihenfolge, um die neue Spalte korrekt einzufügen
spalten = ['MA-ID', 'Funktion', 'Datum', 'Wochentag', 'Dienst',
           'Beliebte Dienste', 'Wochentage beliebte Dienste', 'Erfüllung beliebte Dienste',
           'Unbeliebte Dienste', 'Wochentage unbeliebte Dienste']
df_abgleich = df_abgleich[spalten]


#____________________________Überprüfung unbeliebter Dienste_________________________________

def pruefe_erfuellung_unbeliebter_dienste(row, dienst_mapping):
    unbeliebte_dienste = row['Unbeliebte Dienste']
    wochentage_unbeliebte_dienste = row['Wochentage unbeliebte Dienste']
    aktueller_dienst = row['Dienst']
    aktueller_wochentag = row['Wochentag']

    if pd.isna(unbeliebte_dienste) or unbeliebte_dienste.strip() == '':
        return "keine unbeliebten Dienste eingetragen"

    if aktueller_dienst in dienst_mapping:
        dienst_kategorie = dienst_mapping[aktueller_dienst]
        unbeliebte_dienste_kategorien = {dienst_mapping[dt]: wochentage for dt, wochentage in
                                         zip(unbeliebte_dienste.split(', '), wochentage_unbeliebte_dienste.split('; '))
                                         if dt in dienst_mapping}

        if dienst_kategorie in unbeliebte_dienste_kategorien:
            wochentage = unbeliebte_dienste_kategorien[dienst_kategorie]
            if  wochentage.strip() or aktueller_wochentag in wochentage:
                return False
    return True

# Anwenden der Funktion und Erstellung der Spalte "Beachtung unbeliebte Dienste"
df_abgleich['Beachtung unbeliebte Dienste'] = df_abgleich.apply(
    pruefe_erfuellung_unbeliebter_dienste, axis=1, dienst_mapping=neues_dienst_mapping
)


# Aktualisieren der Spaltenreihenfolge, um die neue Spalte korrekt einzufügen
spalten = ['MA-ID', 'Funktion', 'Datum', 'Wochentag', 'Dienst',
           'Beliebte Dienste', 'Wochentage beliebte Dienste', 'Erfüllung beliebte Dienste',
           'Unbeliebte Dienste', 'Wochentage unbeliebte Dienste', 'Beachtung unbeliebte Dienste']
df_abgleich = df_abgleich[spalten]

# Spalten hinzufügen
df_abgleich.insert(df_abgleich.columns.get_loc('Dienst') + 1, 'Diff', df['Diff'])
df_abgleich.insert(df_abgleich.columns.get_loc('Dienst') + 2, 'Arbeitsblocklänge', df['Arbeitsblocklänge'])
df_abgleich.insert(df_abgleich.columns.get_loc('Dienst') + 3, 'Arbeitsblockdienst', df['Arbeitsblockdienst'])
# Spalten hinzufügen
df_abgleich.insert(df_abgleich.columns.get_loc('Arbeitsblockdienst') + 1, 'Arbeitsblocklänge Opt high', '')
df_abgleich.insert(df_abgleich.columns.get_loc('Arbeitsblocklänge Opt high') + 1, 'Arbeitsblocklänge Opt high erfüllt',
                   '')
df_abgleich.insert(df_abgleich.columns.get_loc('Arbeitsblocklänge Opt high erfüllt') + 1, 'Arbeitsblocklänge Opt low',
                   '')
df_abgleich.insert(df_abgleich.columns.get_loc('Arbeitsblocklänge Opt low') + 1, 'Arbeitsblocklänge Opt low erfüllt',
                   '')
df_abgleich.insert(df_abgleich.columns.get_loc('Arbeitsblocklänge Opt low erfüllt') + 1,
                   'Opt. Freiblockdauer (in Std.)', '')
df_abgleich.insert(df_abgleich.columns.get_loc('Opt. Freiblockdauer (in Std.)') + 1,
                   'Opt. Freiblockdauer (in Std.) erfüllt', '')
df_abgleich.insert(df_abgleich.columns.get_loc('Opt. Freiblockdauer (in Std.) erfüllt') + 1,
                   'Opt. Anz. Dienste in Block', '')
df_abgleich.insert(df_abgleich.columns.get_loc('Opt. Anz. Dienste in Block') + 1, 'Opt. Anz. Dienste in Block erfüllt',
                   '')
df_abgleich.insert(df_abgleich.columns.get_loc('Opt. Anz. Dienste in Block erfüllt') + 1, 'Bestrafte Abfolgen', '')
df_abgleich.insert(df_abgleich.columns.get_loc('Bestrafte Abfolgen') + 1, 'Bestrafte Abfolgen beachtet', '')
df_abgleich.insert(df_abgleich.columns.get_loc('Bestrafte Abfolgen beachtet') + 1, 'Opt. freie Weekends', '')
df_abgleich.insert(df_abgleich.columns.get_loc('Opt. freie Weekends') + 1, 'Opt. freie Weekends erfüllt', '')
df_abgleich.insert(df_abgleich.columns.get_loc('Opt. freie Weekends erfüllt') + 1, 'Opt aufeinanderf. WE mit Arbeit',
                   '')
df_abgleich.insert(df_abgleich.columns.get_loc('Opt aufeinanderf. WE mit Arbeit') + 1,
                   'Opt aufeinanderf. WE mit Arbeit beachtet', '')

# Aktualisieren der Spaltenreihenfolge
spalten = ['MA-ID', 'Funktion', 'Datum', 'Wochentag', 'Dienst', 'Diff', 'Arbeitsblocklänge', 'Arbeitsblockdienst',
           'Arbeitsblocklänge Opt high',
           'Arbeitsblocklänge Opt high erfüllt', 'Arbeitsblocklänge Opt low',
           'Arbeitsblocklänge Opt low erfüllt', 'Opt. Freiblockdauer (in Std.)',
           'Opt. Freiblockdauer (in Std.) erfüllt', "Opt. Anz. Dienste in Block",
           'Opt. Anz. Dienste in Block erfüllt', 'Bestrafte Abfolgen', 'Bestrafte Abfolgen beachtet',
           'Opt. freie Weekends', 'Opt. freie Weekends erfüllt',
           'Opt aufeinanderf. WE mit Arbeit', 'Opt aufeinanderf. WE mit Arbeit beachtet',
           'Beliebte Dienste', 'Wochentage beliebte Dienste', 'Erfüllung beliebte Dienste',
           'Unbeliebte Dienste', 'Wochentage unbeliebte Dienste', 'Beachtung unbeliebte Dienste']
df_abgleich = df_abgleich[spalten]

# Iteriere über jede Zeile im DataFrame df_abgleich
for index, row in df_abgleich.iterrows():
    # Extrahiere die MA-ID aus der aktuellen Zeile
    ma_id = row['MA-ID']

    # Versuche, den entsprechenden Eintrag aus df_excel zu finden
    relevant_data = df_excel.loc[df_excel['ID'] == ma_id, 'Max. + Opt high + Opt-low Arbeitsblocklänge']

    # Verwende next() mit iter(), um sicherzustellen, dass wir keinen IndexError bekommen, wenn keine Daten gefunden werden
    data_entry = next(iter(relevant_data), '')

    # Prüfe, ob die Daten das erwartete Format haben und führe die Aufteilung durch
    if '-' in data_entry:
        parts = data_entry.split('-')
        if len(parts) >= 3:
            opt_high = parts[1].strip()
            opt_low = parts[2].strip()
            df_abgleich.at[index, 'Arbeitsblocklänge Opt high'] = opt_high
            df_abgleich.at[index, 'Arbeitsblocklänge Opt low'] = opt_low
        else:
            df_abgleich.at[index, 'Arbeitsblocklänge Opt high'] = ''
            df_abgleich.at[index, 'Arbeitsblocklänge Opt low'] = ''
    else:
        df_abgleich.at[index, 'Arbeitsblocklänge Opt high'] = ''
        df_abgleich.at[index, 'Arbeitsblocklänge Opt low'] = ''


# Iteriere über jede Zeile im DataFrame df_abgleich
for index, row in df_abgleich.iterrows():
    # Extrahiere die MA-ID aus der aktuellen Zeile
    ma_id = row['MA-ID']

    # Extrahiere Daten und vermeide IndexError
    freiblockdauer_opt_entry = next(iter(df_excel.loc[df_excel['ID'] == ma_id, 'Min. + Opt. Freiblockdauer (in Std.)']), '')
    if '-' in freiblockdauer_opt_entry:
        freiblockdauer_opt = freiblockdauer_opt_entry.split('-')[1].strip()
    else:
        freiblockdauer_opt = ''
    df_abgleich.at[index, 'Opt. Freiblockdauer (in Std.)'] = freiblockdauer_opt

# Iteriere über jede Zeile im DataFrame df_abgleich
for index, row in df_abgleich.iterrows():
    # Extrahiere die MA-ID aus der aktuellen Zeile
    ma_id = row['MA-ID']

    # Extrahiere Daten und vermeide IndexError
    freie_weekends_opt_entry = next(iter(df_excel.loc[df_excel['ID'] == ma_id, 'Min + Opt. freie Weekends']), '')
    if '-' in freie_weekends_opt_entry:
        freie_weekends_opt = freie_weekends_opt_entry.split('-')[1].strip()
    else:
        freie_weekends_opt = ''
    df_abgleich.at[index, 'Opt. freie Weekends'] = freie_weekends_opt

# Iteriere über jede Zeile im DataFrame df_abgleich
for index, row in df_abgleich.iterrows():
    # Extrahiere die MA-ID aus der aktuellen Zeile
    ma_id = row['MA-ID']

    # Extrahiere Daten und vermeide IndexError
    aufeinanderf_we_opt_entry = next(iter(df_excel.loc[df_excel['ID'] == ma_id, 'Max + Opt aufeinanderf. WE mit Arbeit']), '')
    if isinstance(aufeinanderf_we_opt_entry, str) and '-' in aufeinanderf_we_opt_entry:
        aufeinanderf_we_opt = aufeinanderf_we_opt_entry.split('-')[1].strip()
    else:
        aufeinanderf_we_opt = ''
    df_abgleich.at[index, 'Opt aufeinanderf. WE mit Arbeit'] = aufeinanderf_we_opt


#____________________________Überprüfung opt. Anzahl Dienste in Block_________________________________

# Funktion zum Extrahieren der Werte aus der Excel-Spalte
def extrahiere_opt_anz_dienste(wert):
    if not wert:  # Frühe Rückgabe, wenn kein Wert vorhanden ist
        return ''
    einzelwerte = [eintrag.strip() for eintrag in wert.split(',')]
    opt_anz_dienste = []
    for eintrag in einzelwerte:
        teile = eintrag.split(':')
        if len(teile) == 2:
            buchstabe = teile[0].strip()
            wert = teile[1].split('-')[1].strip() if '-' in teile[1] else ''
            opt_anz_dienste.append(f"{buchstabe}: {wert}")
    return ', '.join(opt_anz_dienste)

# Iteriere über jede Zeile im DataFrame df_abgleich
for index, row in df_abgleich.iterrows():
    ma_id = row['MA-ID']
    opt_anz_dienste_entry = next(iter(df_excel.loc[df_excel['ID'] == ma_id, 'Max. + Opt. Anz. Dienste in Block']), '')
    opt_anz_dienste = extrahiere_opt_anz_dienste(opt_anz_dienste_entry)
    df_abgleich.at[index, 'Opt. Anz. Dienste in Block'] = opt_anz_dienste





#____________________________Überprüfung Arbeitsblocklänge_________________________________

df_abgleich['Gruppen-ID'] = (df_abgleich['Arbeitsblocklänge'].isnull() | df_abgleich['Arbeitsblocklänge'].shift(1).isnull()).cumsum()

# Die Summe der Arbeitsblocklängen für jede Gruppe berechnen
df_abgleich['Arbeitsblocklänge Summe'] = df_abgleich.groupby('Gruppen-ID')['Arbeitsblocklänge'].transform('sum')

# Entfernen der Spalte 'Arbeitsblocklänge Opt high erfüllt'
df_abgleich.drop('Arbeitsblocklänge Opt high erfüllt', axis=1, inplace=True)

# Umbenennen der Spalte 'Arbeitsblocklänge Opt low erfüllt' in 'Arbeitsblocklänge Opt high und Opt low erfüllt'
df_abgleich.rename(columns={'Arbeitsblocklänge Opt low erfüllt': 'Arbeitsblocklänge Opt high und Opt low erfüllt'}, inplace=True)

# Funktion zum Überprüfen von Optimum High und Low
def pruefe_arbeitsblocklaenge_opt_erfuellt(row):
    try:
        # Überprüfen, ob die Summe der Arbeitsblocklängen zwischen Optimum Low und High liegt
        if row['Arbeitsblocklänge Summe'] == 0.0:
            return np.nan
        return (row['Arbeitsblocklänge Summe'] >= int(row['Arbeitsblocklänge Opt low'])) and (row['Arbeitsblocklänge Summe'] <= int(row['Arbeitsblocklänge Opt high']))
    except (ValueError, TypeError):
        # Wenn NaN in 'Arbeitsblocklänge Opt low' oder 'Arbeitsblocklänge Opt high' steht, soll NaN zurückgegeben werden
        return np.nan

# Funktion auf den DataFrame anwenden, um die Spalte 'Arbeitsblocklänge Opt high und Opt low erfüllt' zu füllen
df_abgleich['Arbeitsblocklänge Opt high und Opt low erfüllt'] = df_abgleich.apply(pruefe_arbeitsblocklaenge_opt_erfuellt, axis=1)




# Spalten entfernen
df_abgleich = df_abgleich.drop(columns=[
    'Opt. Freiblockdauer (in Std.)',
    'Opt. Freiblockdauer (in Std.) erfüllt',
    'Bestrafte Abfolgen',
    'Bestrafte Abfolgen beachtet'
])

# Neue Spalte "Weekend" hinzufügen, die prüft, ob der Wochentag ein Samstag oder Sonntag ist
df_abgleich['Weekend'] = df_abgleich['Wochentag'].apply(lambda x: x in ['Saturday', 'Sunday'])

# Spalte "Weekend" rechts neben der Spalte "Wochentag" positionieren
wochentag_index = df_abgleich.columns.get_loc('Wochentag') + 1
columns = df_abgleich.columns.tolist()
columns = columns[:wochentag_index] + ['Weekend'] + columns[wochentag_index:-1]

df_abgleich = df_abgleich[columns]

#____________________________Überprüfung Wochenenden mit Arbeit_________________________________

# Zählen, wie viele Wochenenden jeder Mitarbeiter arbeitet
weekends_mit_arbeit = df_abgleich.groupby('MA-ID')['Weekend'].sum()

# Ergebnis dem DataFrame hinzufügen, basierend auf der MA-ID
df_abgleich = df_abgleich.merge(weekends_mit_arbeit.rename('Weekends mit Arbeit'), on='MA-ID', how='left')

# Spalte "Weekends mit Arbeit" rechts neben der Spalte "Opt. freie Weekends" positionieren
opt_freie_weekends_index = df_abgleich.columns.get_loc('Opt. freie Weekends') + 1
columns = df_abgleich.columns.tolist()
columns = columns[:opt_freie_weekends_index] + ['Weekends mit Arbeit'] + columns[opt_freie_weekends_index:-1]

df_abgleich = df_abgleich[columns]

# Das erste und das letzte Datum im gesamten Zeitraum finden
erstes_datum = df_abgleich['Datum'].min()
letztes_datum = df_abgleich['Datum'].max()

# Eine Liste aller Daten zwischen dem ersten und dem letzten Datum erstellen
gesamter_zeitraum = pd.date_range(start=erstes_datum, end=letztes_datum)

# Anzahl der Samstage und Sonntage im gesamten Zeitraum zählen
anzahl_samstage = sum(gesamter_zeitraum.weekday == 5)
anzahl_sonntage = sum(gesamter_zeitraum.weekday == 6)

# Die Anzahl der Samstage und Sonntage im gesamten Zeitraum jeder Zeile hinzufügen
df_abgleich['Anzahl Samstage gesamter Zeitraum'] = anzahl_samstage
df_abgleich['Anzahl Sonntage gesamter Zeitraum'] = anzahl_sonntage

# Berechne die neue Spalte 'Anzahl Wochenende gesamter Zeitraum' als Summe der beiden vorherigen Spalten
df_abgleich['Anzahl Wochenende gesamter Zeitraum'] = df_abgleich['Anzahl Samstage gesamter Zeitraum'] + df_abgleich['Anzahl Sonntage gesamter Zeitraum']


# Spalten "Anzahl Samstage gesamter Zeitraum" und "Anzahl Sonntage gesamter Zeitraum" rechts neben der Spalte "Weekend" positionieren
weekend_index = df_abgleich.columns.get_loc('Weekend') + 1
columns = df_abgleich.columns.tolist()
columns = columns[:weekend_index] + ['Anzahl Samstage gesamter Zeitraum', 'Anzahl Sonntage gesamter Zeitraum'] + columns[weekend_index:]

df_abgleich = df_abgleich[columns]

#____________________________Überprüfung Wochenenden ohne Arbeit_________________________________

# Berechne 'Weekends ohne Arbeit' als Differenz der 'Anzahl Wochenende gesamter Zeitraum' und 'Weekends mit Arbeit'
df_abgleich['Weekends ohne Arbeit'] = df_abgleich['Anzahl Wochenende gesamter Zeitraum'] - df_abgleich['Weekends mit Arbeit']


# Konvertiere 'Opt. freie Weekends' und 'Weekends ohne Arbeit' zu numerischen Typen
df_abgleich['Opt. freie Weekends'] = pd.to_numeric(df_abgleich['Opt. freie Weekends'], errors='coerce')
df_abgleich['Weekends ohne Arbeit'] = pd.to_numeric(df_abgleich['Weekends ohne Arbeit'], errors='coerce')

# Überprüfe, ob 'Weekends ohne Arbeit' größer oder gleich 'Opt. freie Weekends' ist
df_abgleich['Opt. freie Weekends erfüllt'] = df_abgleich['Weekends ohne Arbeit'] >= df_abgleich['Opt. freie Weekends']

# Markiere Wochenenden in einer neuen Spalte
df_abgleich['Ist Wochenende'] = df_abgleich['Wochentag'].isin(['Saturday', 'Sunday'])

# Erstelle eine Gruppen-ID für aufeinanderfolgende Zeilen (für jeden Mitarbeiter separat)
df_abgleich['Gruppen-ID'] = ((df_abgleich['MA-ID'] != df_abgleich['MA-ID'].shift()) |
                             (df_abgleich['Ist Wochenende'] != df_abgleich['Ist Wochenende'].shift()) |
                             (~df_abgleich['Ist Wochenende'])).cumsum()

#____________________________Überprüfung aufeinanderfolgenden Wochenenden mit Arbeit_________________________________

# Berechne aufeinanderfolgende Wochenenden
df_abgleich['Aufeinanderfolgende WE'] = df_abgleich.groupby('Gruppen-ID')['Ist Wochenende'].transform('sum')

# Setze Werte auf 0, wo 'Ist Wochenende' False ist, um nur die tatsächlichen aufeinanderfolgenden Wochenenden zu zählen
df_abgleich['Aufeinanderfolgende WE mit Arbeit'] = df_abgleich.apply(lambda x: x['Aufeinanderfolgende WE'] if x['Ist Wochenende'] else 0, axis=1)

# Berechne die Anzahl aufeinanderfolgender Wochenenden pro MA-ID
df_abgleich['Aufeinanderfolgende WE mit Arbeit'] = df_abgleich.groupby('MA-ID')['Aufeinanderfolgende WE mit Arbeit'].transform('max')

# Füge die Spalte 'Aufeinanderfolgende WE mit Arbeit' rechts neben 'Opt aufeinanderf. WE mit Arbeit' ein
opt_aufeinanderf_we_mit_arbeit_index = df_abgleich.columns.get_loc('Opt aufeinanderf. WE mit Arbeit') + 1
df_abgleich.insert(opt_aufeinanderf_we_mit_arbeit_index, 'Aufeinanderfolgende WE mit Arbeit', df_abgleich.pop('Aufeinanderfolgende WE mit Arbeit'))

# Überprüfung und Eintrag in 'Opt aufeinanderf. WE mit Arbeit beachtet'
df_abgleich['Opt aufeinanderf. WE mit Arbeit'] = pd.to_numeric(df_abgleich['Opt aufeinanderf. WE mit Arbeit'], errors='coerce')
df_abgleich['Opt aufeinanderf. WE mit Arbeit beachtet'] = df_abgleich['Aufeinanderfolgende WE mit Arbeit'] <= df_abgleich['Opt aufeinanderf. WE mit Arbeit']

# Aufräumen: Entferne Hilfsspalten
df_abgleich.drop(['Ist Wochenende', 'Gruppen-ID', 'Aufeinanderfolgende WE'], axis=1, inplace=True, errors='ignore')

#____________________________Überprüfung Anzahl Dienste in Block_________________________________


# Generierung einer Block ID für jeden aufeinanderfolgenden Block
df_abgleich['Block-ID'] = ((df_abgleich['Arbeitsblocklänge'] != df_abgleich['Arbeitsblocklänge'].shift(1)) | (df_abgleich['MA-ID'] != df_abgleich['MA-ID'].shift(1))).cumsum()

# Filtern, um Zeilen beizubehalten, in denen „Arbeitsblocklänge“ grösser als 0 ist, was auf einen aktiven Arbeitsblock hinweist
active_blocks = df_abgleich[df_abgleich['Arbeitsblocklänge'] > 0]

# Gruppieren  nach „MA-ID“ und „Block-ID“, um das Vorkommen jedes Dienstes innerhalb eines Blocks zu zählen
service_counts_per_block = active_blocks.groupby(['MA-ID', 'Block-ID', 'Arbeitsblockdienst']).size().reset_index(name='Count')

# Erstellen einer neuen Spalte, die Dienst und Anzahl im gewünschten Format kombiniert (z. B. „F: 3“)
service_counts_per_block['Service:Count'] = service_counts_per_block['Arbeitsblockdienst'] + ': ' + service_counts_per_block['Count'].astype(str)

# Zusammenfassen des „Service:Count“-Strings wieder zu einem einzigen String pro Block
block_summary = service_counts_per_block.groupby(['MA-ID', 'Block-ID'])['Service:Count'].apply(lambda x: ', '.join(x)).reset_index()

# Merge der Zusammenfassung wieder in das ursprüngliche DataFrame
df_abgleich = df_abgleich.merge(block_summary, on=['MA-ID', 'Block-ID'], how='left')

# Löschen der "Block-ID", da sie nicht mehr benötigt wird
df_abgleich.drop('Block-ID', axis=1, inplace=True)

# Umbennung von 'Service:Count' zu 'Anzahl Dienste pro Block'
df_abgleich.rename(columns={'Service:Count': 'Anzahl Dienste pro Block'}, inplace=True)

# Positionierung der Spalte „Anzahl Dienste pro Block“ direkt neben „Opt. Anz. Dienste in Block'
opt_dienste_index = df_abgleich.columns.get_loc('Opt. Anz. Dienste in Block') + 1

# Spalten neu anordnen, um „Anzahl Dienste pro Block“ richtig zu positionieren
columns = df_abgleich.columns.tolist()
columns.insert(opt_dienste_index, columns.pop(columns.index('Anzahl Dienste pro Block')))
df_abgleich = df_abgleich[columns]



def pruefe_erfuellung(row, dienst_mapping):
    if pd.isna(row['Opt. Anz. Dienste in Block']) or row['Opt. Anz. Dienste in Block'].strip() == '':
        return np.nan  # Rückgabe von NaN, wenn keine Daten vorliegen

    # Sichere Umwandlung in Wörterbücher unter Verwendung von get() und Kontrolle der Wertformatierung
    opt_anz_dict = {}
    for d in row['Opt. Anz. Dienste in Block'].split(','):
        parts = d.split(':')
        if len(parts) == 2:
            key = parts[0].strip()
            # TUGI: int() hinzugefügt => str zu int
            value = int(parts[1].strip())
            # TUGI: warum .get(key) und weshalb .isdigit()?
            if key in dienst_mapping.values():
            # if dienst_mapping.get(key) and value.isdigit():
                opt_anz_dict[key] = int(value)

    if pd.isna(row['Anzahl Dienste pro Block']) or row['Anzahl Dienste pro Block'].strip() == '':
        return np.nan  # Rückgabe von NaN, wenn keine Daten vorliegen

    anz_block_dict = {}
    for d in row['Anzahl Dienste pro Block'].split(','):
        parts = d.split(':')
        if len(parts) == 2:
            key = parts[0].strip()
            value = int(parts[1].strip())
            # TUGI: ergänzt
            if key not in dienst_mapping.keys():
                raise ValueError(f"Dienst {key} nicht in dienst_mapping...")
            # TUGI: Übersetzen
            key = dienst_mapping[key]
            if key in dienst_mapping.values():
                if key in anz_block_dict.keys():
                    anz_block_dict[key] = anz_block_dict[key] + int(value)
                else:
                    anz_block_dict[key] = int(value)

    # Überprüfung, ob mindestens ein Wert übereinstimmt
    for kategorie in opt_anz_dict:
        if kategorie in anz_block_dict and anz_block_dict[kategorie] <= opt_anz_dict[kategorie]:
            return True  # Erfüllt, wenn mindestens ein Wert übereinstimmt oder kleiner ist
    return False  # Nicht erfüllt, wenn kein Wert übereinstimmt oder kleiner ist

# Anwenden der Funktion auf df_abgleich
df_abgleich['Opt. Anz. Dienste in Block erfüllt'] = df_abgleich.apply(
    pruefe_erfuellung, dienst_mapping=neues_dienst_mapping, axis=1
)

#_______________________________________________________________________
# # Pfad, unter dem die Excel-Datei gespeichert werden soll
# excel_dateipfad = 'df_abgleich.xlsx'
#
# # Exportieren des DataFrames als Excel-Datei
# df_abgleich.to_excel(excel_dateipfad, index=False, engine='openpyxl')

breakpoint()
#_______________________________________________________________________
# Neues DF für Auswertung erstellen
neues_df = df_abgleich[['MA-ID', 'Datum', 'Arbeitsblocklänge Opt high und Opt low erfüllt',  'Opt. Anz. Dienste in Block', 'Anzahl Dienste pro Block', 'Opt. Anz. Dienste in Block erfüllt',
                        'Opt. freie Weekends erfüllt', 'Opt aufeinanderf. WE mit Arbeit beachtet',
                        'Erfüllung beliebte Dienste', 'Beachtung unbeliebte Dienste']].copy()

# Oder als Excel-Datei, wenn du Formatierungen oder spezielle Excel-Features erhalten möchtest
neues_df.to_excel('df_Januar 2024.xlsx', index=False)



breakpoint()

# Demonstration der Ergebnisse
print("Daten aus JSON-Datei:")
print(df.head())

print("\nDaten aus Excel-Datei:")
print(df_excel.head())

print("\nAngepasste Daten mit Erfüllung der beliebten Dienste:")
print(df_abgleich[['MA-ID', 'Beliebte Dienste', 'Wochentage beliebte Dienste', 'Erfüllung beliebte Dienste']].head())






















