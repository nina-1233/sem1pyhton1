# Further readings: https://www.datacamp.com/community/tutorials/python-excel-tutorial
# WICHTIG: Neben pandas und matplotlib muss zusätzlich noch die openpyxl Bibliothek 
# imstalliert werden. Diese wird im Hintergrund für das Öffnen der Exceldatei in der
# Anweisung pd.ExcelFile(in_file) in xfile_read() benötigt.

import pandas as pd
import matplotlib.pyplot as plt

DEBUG_INFO = True

def xfile_read(wdir, in_file, tabble_name):

    # richtigen Pfad zusammensetzen
    in_file = wdir + in_file
    print(f"Loading file '{in_file}' ...")

    # erstellt ein Excel-File in pandas 
    xl = pd.ExcelFile(in_file)
    if DEBUG_INFO: print("All sheet names: ", xl.sheet_names)

    # table_name richtig benennen
    if tabble_name == "": tabble_name = 'Quelldaten'
    print(f"Parsing data from sheet '{tabble_name}'")

    # Data Frame aus Excel-File erzeugen
    df = pd.read_excel(in_file)

    # Namen der Spalten werden ausgegeben
    header_row = df.columns.tolist()
    print(header_row)

    # Anazhl der Spalten
    number_of_headers= len(header_row)
    print(number_of_headers)

    return df, header_row, number_of_headers

'''
plot_client_data: plotting dataframes slices using 'Kunde' or 'Produkt' and 'Quartal' as slicing params 
input:
- df: data frame 
- kunde_produkt: 'Kunde' or 'Produkt' (slicing param) 
- quartal: 'Qrtl x' or 'all' (slicing param)
output: 
- None (the procedures will use plt.show() internally to plot a graph)
'''
def plot_client_data(df, kunde_produkt, quartal):

    # abhängig von Auswahl werden die passenden Daten (Kundenname oder Produktbezeichnung) aus Data Frame entnommen
    if kunde_produkt == 'Kunde':
        x_axis_raw = df['Kunde']
    if kunde_produkt == 'Produkt':
        x_axis_raw = df['Produkt']
    
    # wird in Liste umgewandelt und anschließend bereinigt von Duplikaten 
    Liste = list(x_axis_raw)
    x_axis_clean = set(Liste)
 
    if DEBUG_INFO: print(f"x_axis_clean ('{kunde_produkt}'):\n{x_axis_clean}")

    # leeres Dictionary wird inalisiert 
    y_axis = {}
    # die Werte des Dictionary werden gleich null gesetzt 
    for x_axis_value in x_axis_clean:
        y_axis[x_axis_value] = 0

    # Summe der ausgewählten Variable (Kunde oder Produkt) wird für einzelne Quartale oder für das ganze Jahr gebildet           
    for x_axis_value in x_axis_clean:
        # Besonderheit, da man 'all' nicht als Variable benutzen kann
        if quartal == 'all':    
            y_axis[x_axis_value] = df.loc[df[kunde_produkt] == x_axis_value, 'Qrtl 1'].sum() + \
            df.loc[df[kunde_produkt] == x_axis_value, 'Qrtl 2'].sum() + \
            df.loc[df[kunde_produkt] == x_axis_value, 'Qrtl 3'].sum() + \
            df.loc[df[kunde_produkt] == x_axis_value, 'Qrtl 4'].sum()
        else: 
            y_axis[x_axis_value] = df.loc[df[kunde_produkt] == x_axis_value, quartal].sum()

    # Ausgabe der ausgewählten Variable mit den passenden Summe im jeweiligen ausgewählten Quartal
    print(f"\nSumme der Umsätze über das Quartal '{quartal}' für '{kunde_produkt}' ...\n{y_axis}")

    # Gesamtumsatz der Datei
    gesamtumsatz = df['Qrtl 1'].sum() + df['Qrtl 2'].sum() + df['Qrtl 3'].sum() + df['Qrtl 4'].sum()

    # Berechnung der Summe des ausgewählten Quartals
    Summe_Quartal = 0
    for wert in y_axis.values():
        Summe_Quartal += wert
 
    # Berechnung des prozentualen Anteils jedes Kunden oder Produkts am Gesamtumsatz
    prozentualer_anteil= (Summe_Quartal/gesamtumsatz)*100
 
    #Ausagbe des prozentualen Anteils des Quartals in Vergleich zum Gesamtumsatz
    print(f"\nProzentualer Anteil am Gesamtumsatzes:\n{prozentualer_anteil}")

    # Optionale Aufgabe 4 (prozentuale Angabe)
    # Schleife um alle Werte von y_axis zur Verrechnung mit dem ausgewählten Quartal um für jeden Kunden / Produkt einen Prozentsatz auzugegen
    for x_axis_value in y_axis.values():
        # Besonderheit bei all, deswegen extra Fall 
        if quartal == 'all':
            # Prozentualer Anteil wird berechnet im Vergleich zum Gesamtzumsatz
            prozent = (x_axis_value / gesamtumsatz) * 100  
            # der jeweilige Schlüssel wird dazu ermittelt
            gesuchter_schluessel = list(y_axis.keys())[list(y_axis.values()).index(x_axis_value)]
            # Ausgabe
            print(f"\nProzentualer Anteil der Umsätze am Gesamtumsatz über das Quartal '{quartal}' für '{gesuchter_schluessel}' ...\n{prozent}") 
        else:
            # Prozentualer Anteil wird berechnet vom jeweiligen ausgewählten Quartal
            prozent = (x_axis_value / df[quartal].sum()) * 100  
            # der jeweilige Schlüssel wird dazu ermittelt
            gesuchter_schluessel = list(y_axis.keys())[list(y_axis.values()).index(x_axis_value)]
            # Ausgabe
            print(f"\nProzentualer Anteil der Umsätze am Gesamtumsatz über das Quartal '{quartal}' für '{gesuchter_schluessel}' ...\n{prozent}") 
    
    # x-Achse: Personen aus den Schlüsselwerten, y-Achse: Ergebnisse aus den Werten zu den Schlüsseln
    Personen = list(y_axis.keys())
    Ergebnisse = list (y_axis.values())

    # Punktediagramm
    plt.scatter(Personen, Ergebnisse, color='red', marker='o') 
    # marker => In welcher Form die Punkte angegeben werden sollen, erst die x-Achse dann die y-Achse
    
    plt.title (f'Umsätze nach {kunde_produkt} im Quartal {quartal}') #Titel des Diagramms
 
    #Beschriftung der Achsen
    plt.xlabel(kunde_produkt)
    plt.xticks(rotation='vertical')
    plt.ylabel(quartal)
 
    plt.grid(True) #Gitterlinien für eine besser übersicht
    plt.show()
 
    #Säulensiagramm
    plt.bar(Personen, Ergebnisse, color='blue', alpha=1) #alpha => Tranzparenz der Säulen, auch hier erst x_Achse dann y-Achse
    plt.title(f'Umsätze nach {kunde_produkt} im Quartal {quartal}') #Titel des Diagramms
 
    #Beschrfitung der Achsen
    plt.xlabel(kunde_produkt)
    plt.xticks(rotation='vertical')
    plt.ylabel(quartal)
 
    plt.grid(axis='y') #Auf wechler Achse die Säulen beginnen sollen
    plt.show()
 
    '''
    # You can specify a rotation for the tick labels in degrees or with keywords.
    # https://matplotlib.org/gallery/ticks_and_spines/ticklabels_rotation.html#sphx-glr-gallery-ticks-and-spines-ticklabels-rotation-py
    plt.xticks(rotation='vertical')
    # Tweak spacing to prevent clipping of tick-labels
    plt.subplots_adjust(bottom=0.4)
    # adjusting axis intervalls if needed
    # plt.axis([0, 6, 0, 20])
    # labelling of axes
    plt.ylabel(quartal)
    # display plot
    plt.show()
    '''

    return 


# MAIN
if __name__ == "__main__":

    # Initialize working directory, file and tab name
    wdir = r'C:\GIT\sem1pyhton1' 
    infile = r'\101_Umsatzbericht.xlsx'
    table_name = 'Quelldaten'
    try:
        df, header_row, number_of_headers = xfile_read(wdir, infile, table_name)
    except:
        print("Sorry. Could not parse input data correctly. Did you complete your work inside xfile_read()?")
        raise SystemExit

    # Simple toggle to check for correct excel file parsing
    parsing_completed = False

    while not parsing_completed:
        parsing_completed = True

        print("Kopfzeile: ", end="")
        print(header_row)
        print("Was möchten Sie analysieren? Kunde (K) oder Produkt (P)?")
        kunde_produkt = input(">>> ")

        print("Welches Quartal möchten Sie betrachten? (1, 2, 3, 4, alle)?")
        quartal = input(">>> ")

        if kunde_produkt == "K" or kunde_produkt == "k": 
            kunde_produkt = "Kunde"
        elif kunde_produkt == "P" or kunde_produkt == "p": 
            kunde_produkt = "Produkt"
        elif kunde_produkt == "":
            break
        else: 
            parsing_completed = False

        if quartal == "1": 
            quartal = "Qrtl 1"
        elif quartal == "2": 
            quartal = "Qrtl 2"
        elif quartal == "3": 
            quartal = "Qrtl 3"
        elif quartal == "4": 
            quartal = "Qrtl 4"
        elif quartal == "alle" or quartal == "all" or quartal == "a":
            quartal = "all"
        elif quartal == "":
            break
        else:
            parsing_completed = False

        # we can only continue plotting the data, if parsing was completed successfully
        if parsing_completed: 
            plot_client_data(df, kunde_produkt, quartal)
        else:
            print("Sorry, no such data to be analyzed. Exiting.")
