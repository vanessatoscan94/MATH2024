import pandas as pd
import os
#
# # Einlesen der Excel-Dateien
# df_januar = pd.read_excel('df_Januar 2023.xlsx')
# df_februar = pd.read_excel('df_Februar 2023.xlsx')
#
# # Zusammenführen der DataFrames
# # Option 1: Vertikales Zusammenführen (Anhängen), wenn sie identische Spalten haben
# df_gesamt = pd.concat([df_januar, df_februar], ignore_index=True)
#
#
#
# # Anzeigen der zusammengefügten Daten
# print(df_gesamt.head())
#
# # Speichern des zusammengefügten DataFrame in einer neuen Excel-Datei
# df_gesamt.to_excel('df_gesamt_2023.xlsx', index=False)


# Beispiel DataFrame
data = pd.read_excel('merged_dataframe.xlsx')

df = pd.DataFrame(data)

# print("Vor der Umwandlung ins Longformat:")
# print(df)

# Umwandlung ins Longformat
long_df = pd.melt(df, id_vars=['MA-ID', 'Datum'], var_name='Präferenz', value_name='Erfüllung')

print("\nNach der Umwandlung ins Longformat:")
print(long_df)

long_df.to_excel('df_merged_long'
                 '.xlsx', index=False)



#------------------Alle DFs einlesen---------------

# Verzeichnis, in dem sich Ihre Dateien befinden
verzeichnis_pfad = 'Dataframes'

# Liste zum Speichern der eingelesenen DataFrames
dfs = []

# Muster für Dateinamen
datei_muster = 'df_*.xlsx'  # Beispiel: Alle Dateien, die mit 'beispiel_' beginnen und die Endung '.xlsx' haben

# Iteriere über alle Dateien im Verzeichnis
for datei in os.listdir(verzeichnis_pfad):
    if datei.endswith('.xlsx') and datei.startswith('df_'):  # Überprüfen Sie, ob die Datei dem Muster entspricht
        datei_pfad = os.path.join(verzeichnis_pfad, datei)
        df = pd.read_excel(datei_pfad)  # DataFrame aus der Excel-Datei einlesen
        dfs.append(df)  # DataFrame zur Liste hinzufügen

# Überprüfen der eingelesenen DataFrames
for idx, df in enumerate(dfs, 1):
    print(f"DataFrame {idx}:")
    print(df)

#------------------Spalte entfernen---------------

# Name der Spalte, die entfernt werden soll
spalte_zu_entfernen = 'Anzahl Dienste pro Block', 'Opt. Anz. Dienste in Block'

# Iteriere über jedes DataFrame und entferne die gewünschte Spalte
for df in dfs:
    if spalte_zu_entfernen in df.columns:
        df.drop(columns=[spalte_zu_entfernen], inplace=True)

# Überprüfen der DataFrames nach Entfernen der Spalte
for idx, df in enumerate(dfs, 1):
    print(f"DataFrame {idx} nach Entfernen der Spalte {spalte_zu_entfernen}:")
    print(df)


#------------------DFs zusammenfügen---------------

# Zusammenführen der DataFrames zu einem einzigen DataFrame
merged_df = pd.concat(dfs, ignore_index=True)

# Speichern des zusammengeführten DataFrames als Excel-Datei
ausgabedatei_pfad = 'merged_dataframe.xlsx'
merged_df.to_excel(ausgabedatei_pfad, index=False)

print("Zusammengeführter DataFrame:")
print(merged_df)

print(f"\nDer zusammengeführte DataFrame wurde erfolgreich als '{ausgabedatei_pfad}' gespeichert.")


