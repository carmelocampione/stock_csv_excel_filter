import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

def verwerk_csv():
    # Selecteer inputbestand
    input_file = filedialog.askopenfilename(
        title="Selecteer CSV-bestand",
        filetypes=[("CSV bestanden", "*.csv")]
    )
    if not input_file:
        return

    # Selecteer output Excel-bestand
    excel_output = filedialog.asksaveasfilename(
        title="Opslaan als Excel",
        defaultextension=".xlsx",
        filetypes=[("Excel bestanden", "*.xlsx")]
    )
    if not excel_output:
        return

    try:
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

        # Wegschrijven naar Excel
        df.to_excel(excel_output, index=False)

        # Bevestiging
        messagebox.showinfo("Succes", f"Excel-bestand opgeslagen als:\n{excel_output}")

    except Exception as e:
        messagebox.showerror("Fout", f"Er is een fout opgetreden:\n{e}")

# Maak GUI
root = tk.Tk()
root.title("CSV naar Excel Verwerker")

frame = tk.Frame(root, padx=20, pady=20)
frame.pack()

label = tk.Label(frame, text="Klik op de knop om een CSV-bestand te verwerken naar Excel.")
label.pack(pady=10)

button = tk.Button(frame, text="Selecteer en verwerk CSV", command=verwerk_csv)
button.pack(pady=10)

root.mainloop()
