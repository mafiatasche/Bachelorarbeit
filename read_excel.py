import pandas as pd

# Lese die Excel-Datei mit erster Zeile als Header
df = pd.read_excel('F:/Friehe_messungen/Vermessung.xlsx', sheet_name=0, header=0)
print('Spalten:', df.columns.tolist())
print('\nErste 5 Zeilen:')
pd.set_option('display.max_columns', None)
pd.set_option('display.width', None)
print(df.head(5).to_string())
print('\nLetzte 5 Zeilen:')
print(df.tail(5).to_string())
