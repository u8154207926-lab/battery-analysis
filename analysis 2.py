#1.Bibliotheken importieren
import pandas as pd #pandas (pd) für Tabellen (Excel lesen, bearbeiten)
import matplotlib.pyplot as plt #Matplotlib (plt) für Diagramme
import os #für Datei- und Ordneroperationen (Pfad, Existenz prüfen, etc.)

#2. Ordner festlegen
# Eingabe- und Ausgabeordner festlegen
input_folder = r"C:\Enrico\New\Python\23.04.2026 2\Batteries analysis\data" #wo die Excel-Datei liegt
output_folder = r"C:\Enrico\New\Python\23.04.2026 2\Batteries analysis\Output" #wo Ergebnisse gespeichert werden ;r"..." = raw string → wichtig für \ in Windows

#3. Output-Ordner erstellen, Bedeutet: Erstelle den Ordner, falls er nicht existiert
exist_ok=True #→ kein Fehler, wenn er schon da ist
os.makedirs(output_folder, exist_ok=True)

#4. Benutzer gibt Dateinamen ein, Programm wartet auf deine Eingabe z. B.: battery_data.xlsx
# Nur Dateiname eingeben (ohne Pfad)
file_name = input("Enter file name (e.g. battery_data.xlsx): ")

#5. Vollständigen Pfad bauen ,Kombiniert: Ordner + Dateiname Ergebnis: C:\...\data\battery_data.xlsx Vorteil: funktioniert sicher auf jedem System
input_path = os.path.join(input_folder, file_name)

#6. Prüfen ob Datei existiert
if not os.path.exists(input_path): #Wenn Datei nicht existiert: Fehlermeldung Programm stoppt
    print("File not found!")
    exit()

#7. Excel-Datei laden 
df = pd.read_excel(input_path) #Datei wird als Tabelle (DataFrame) geladen

#8. Daten bereinigen
df["Capacity (mAh)"] = pd.to_numeric(df["Capacity (mAh)"], errors="coerce") #Bedeutung:Text → Zahl umwandeln, Fehler → wird zu NaN (leerer Wert)
df["Voltage (V)"] = pd.to_numeric(df["Voltage (V)"], errors="coerce")
df["Temperature (°C)"] = pd.to_numeric(df["Temperature (°C)"], errors="coerce")
df["Resistance (Ohm)"] = pd.to_numeric(df["Resistance (Ohm)"], errors="coerce")

df = df.dropna().reset_index(drop=True) #Bedeutung:dropna() → löscht alle Zeilen mit Fehlern, reset_index() → setzt Index neu (0,1,2,3,...)

#9. Sortieren (FIX!)
df = df.sort_values(by="Cycle") # Sortiert Tabelle nach „Cycle“ ,von kleinen zahlen ornden zu großen zahlen

#10. Neue Spalten berechnen ,#Bedeutung: Vergleich mit erstem Wert zeigt prozentualen Verlust
df["Capacity fade"] = ((df["Capacity (mAh)"][0] - df["Capacity (mAh)"]) / df["Capacity (mAh)"][0]) * 100
df["Capacity Loss"] = df["Capacity (mAh)"].diff() #Bedeutung: Unterschied zum vorherigen Wert
#
#
#11. Dateinamen vorbereiten,  ohne .xlsx holen
base = os.path.splitext(file_name)[0] #Beispiel: battery_data.xlsx → battery_data ; .xlsx wird entfernt 
#
#
#12. Speicherpfade erstellen
output_excel = os.path.join(output_folder, "processed_" + base + ".xlsx") #Ergebnis: processed_battery_data.xlsx und plot_battery_data.png im Output-Ordner gespeichert
output_plot = os.path.join(output_folder, "plot_" + base + ".png")

#13. Excel speichern 
df.to_excel(output_excel, index=False) #Speichert Tabelle als Excel, index=False → keine extra Spalte mit Zahlen

#14. Plot erstellen 
plt.plot(df["Cycle"], df["Capacity (mAh)"]) #x-Achse: Cycle ; y-Achse: Capacity

#15. Plot verschönern ;Achsen + Titel + Raster
plt.xlabel("Cycle") 
plt.ylabel("Capacity (mAh)")
plt.title("Capacity vs Cycle")
plt.grid(True)

#16. Plot speichern
plt.savefig(output_plot) #speichert Diagramm als PNG

#17. Ausgabe im Terminal 
print("✅ Done!") #zeigt: Programm fertig, Speicherorte
print("Saved Excel:", output_excel)
print("Saved Plot:", output_plot)

#Gesamtlogik (ganz einfach)
#Datei auswählen
#Datei laden
#Daten bereinigen
#Daten berechnen
#Ergebnisse speichern
#Diagramm erstellen