import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import os

def verwerk_bestand(pad):
    try:
        df = pd.read_csv(pad, skiprows=4, sep=';', header=None, dtype=str)
        df = df.iloc[:, :8]

        for col in df.columns:
            df[col] = df[col].map(lambda x: x.replace('="', '').replace('"', '').strip() if isinstance(x, str) else x)

        df = df.iloc[:, [0, 2, 3, 4, 6]]
        df.columns = ['locatie', 'UN', 'Product', 'batch', 'aantal']
        df['verdiep'] = df['locatie'].str[-2:]

        bestand_naam = os.path.splitext(os.path.basename(pad))[0]
        map_pad = os.path.dirname(pad)
        csv_output = os.path.join(map_pad, f"{bestand_naam}_bewerkt.csv")
        excel_output = os.path.join(map_pad, f"{bestand_naam}_bewerkt_excel.xlsx")

        df.to_csv(csv_output, index=False, sep=';')
        df.to_excel(excel_output, index=False)

        messagebox.showinfo("Succes", f"Bestanden opgeslagen:\n\nCSV: {csv_output}\nExcel: {excel_output}")
    except Exception as e:
        messagebox.showerror("Fout", f"Er ging iets mis:\n{e}")

def selecteer_bestand():
    bestand = filedialog.askopenfilename(filetypes=[("CSV bestanden", "*.csv")])
    if bestand:
        verwerk_bestand(bestand)

# GUI setup
root = tk.Tk()
root.title("STOCK Verwerker")
root.geometry("400x200")
root.resizable(False, False)

label = tk.Label(root, text="Sleep je STOCK.CSV bestand of klik op de knop", font=("Arial", 12))
label.pack(pady=20)

btn_selecteer = tk.Button(root, text="Selecteer STOCK.CSV", command=selecteer_bestand, width=25, height=2)
btn_selecteer.pack()

root.mainloop()
