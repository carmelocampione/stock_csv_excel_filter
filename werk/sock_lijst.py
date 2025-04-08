import pandas as pd

# Bestandspaden
input_file = 'C:/Users/carcam/PycharmProjects/werk/STOCK.CSV'
output_file = 'C:/Users/carcam/PycharmProjects/werk/STOCK_bewerkt.csv'
excel_output = 'C:/Users/carcam/PycharmProjects/werk/STOCK_bewerkt_excel.xlsx'

# Inlezen van het CSV-bestand (gebruik juiste encoding)
df = pd.read_csv(input_file, skiprows=4, sep=';', header=None, dtype=str, encoding='ISO-8859-1')

# Beperk tot eerste 8 kolommen
df = df.iloc[:, :8]

# Schoonmaken van alle kolommen
for col in df.columns:
    df[col] = df[col].map(lambda x: x.replace('="', '').replace('"', '').strip() if isinstance(x, str) else x)

# Selecteer relevante kolommen
df = df.iloc[:, [0, 2, 3, 4, 6]]
df.columns = ['locatie', 'UN', 'Product', 'batch', 'aantal']

# Voeg een kolom toe met de laatste twee tekens van 'locatie'
df['verdiep'] = df['locatie'].str[-2:]

# Wegschrijven naar CSV en Excel
df.to_csv(output_file, index=False, sep=';')
df.to_excel(excel_output, index=False)

# Print bevestiging
print("Bestanden opgeslagen als:")
print(f"CSV: {output_file}")
print(f"Excel: {excel_output}")


