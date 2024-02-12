'''Further readings: https://www.datacamp.com/community/tutorials/python-excel-tutorial
# WICHTIG: Neben pandas und matplotlib muss zusätzlich noch die openpyxl Bibliothek
# imstalliert werden. Diese wird im Hintergrund für das Öffnen der Exceldatei in der
# Anweisung pd.ExcelFile(in_file) in xfile_read() benötigt.
 
# durch das importieren / installieren der Biblotheken können wir auf diese zugreifen  '''
 
import pandas as pd
import matplotlib.pyplot as plt
 
# Zwischenschritte werden bei False ausgeblendet z.B. seheet_names weden ausgeblendet bei False und bei True eingebelendet
DEBUG_INFO = False
 
# Erste Funktion (Funktionen definiert man, um in der "Main" Zeilen zu sparen, da man meistens mit mehreren Funktionen bzw. mit öfter genutze Funktionen arbeiten kann)
def xfile_read(wdir, in_file, tabble_name):                                                                                                         # Bennenung der Fuktion  mit übergebenen Variablen , wdir (lokaler Speicherort), in_file (Name der Ecxel Datei)), table_name (Überschirft)
 
# Richtigen Pfad zusammensetzen
    in_file = wdir + in_file                                                                                                                        # Speicherort der Excel Datei = Speicherort (wdir) + Exceldatei (infile)
    print(f"Loading file '{in_file}' ...")                                                                                                          # Der Nutzer wird über den Speicherort der Excel Tabelle informiert
 
# Erstellt ein Excel-File in pandas / pandas ist eine programm Biblothek für Phyton für zur Erarbeitungg, Analyse und Darstellung von Daten
    xl = pd.ExcelFile(in_file)                                                                                                                    # Wird ein Objekt pd.Excelfile erstellt / als Argument nehmen wir die 'in_file' variable, damit das Programm weiß weclhes file es einlesen soll
    if DEBUG_INFO: print("All sheet names: ", xl.sheet_names)                                                                                        # if (wenn) sorgt dafür, dass der Python-Code ausgeführt wird, wenn eine Bedigung erfüllt ist / print (gibt aus) welche sheets in dieser Excel file vorhanden ist / Es sind mehrere Tabellen in dieser Datei
 
# table_name richtig benennen
    if tabble_name == "": tabble_name = 'Quelldaten'                                                                                                 # Wenn der Tabellen Name unbennant ist wird dieseer mit 'Quellendaten' überschireben
 
# Gibt die Überschriften der Tabelle 'Quellendaten' aus
    print(f"Parsing data from sheet '{tabble_name}'")
 
# Aufabe 1
    
# Data Frame aus Excel-File erzeugen
    df = pd.read_excel(in_file)                                                                                                                      # Einlesen des files als pandas Dataframe (Duplikat von Excel mit dem Phyton arbeiten kann)
 
# Namen der Spalten werden ausgegeben
    header_row = df.columns.tolist()                                                                                                                 # heade_row (Variable (Kopfzeile)) / df.colums.tolist <Funktion aus der Biblothek> (übergibt die erste Zeile von DataFrame als Liste)
    print(header_row)                                                                                                                                # Aufgerufene Liste aus der vorgherigen Zeile wird ausgegeben
 
# Anazhl der Spalten
    number_of_headers= len(header_row)                                                                                                               # len <Funktion> (übergibt die Anzahl der z.B Spalten in der Kopfzeile)
    print(number_of_headers)                                                                                                                         # Ausgabe der Anzahl Spalten
 
    return df, header_row, number_of_headers                                                                                                         # Varaibeln dieser Funktion werden an die Main zurückgeben (sonst wirft er den Nutzer ins except)
 
 
def plot_client_data(df, kunde_produkt, quartal):
 
# Aufgabe 2 (Sortierung der Tabelle, Verhinderung von Wiederholung)
    
# Abhängig von Auswahl werden die passenden Daten (Kundenname oder Produktbezeichnung) aus Data Frame entnommen
    if kunde_produkt == 'Kunde':                                                                                                                      # Kunden werden ausgelesen und numeriert
        x_axis_raw = df['Kunde']                                                                                                                      # x_axis_raw vorgegebene Variable -> Dictonaray / nach der Auslesung werden diesen Informationen ins Dictonray eingetragen
    elif kunde_produkt == 'Produkt':                                                                                                                  # Produkte werden ausgelesen und nummeriert
        x_axis_raw = df['Produkt']                                                                                                                    # Nach der Auslesung werden diesen Informationen ins Dictonray eingetragen
   
# Wird in Liste umgewandelt und anschließend bereinigt von Duplikaten
    Liste = list(x_axis_raw)                                                                                                                          # Listen wird erstellt aus dem Dictonary
    x_axis_clean = set(Liste)                                                                                                                         # Bereinigung der Duplikate (set: löschen) / z.B. vorher 277 nachher 78 bei Kunde /raw: mit Duplikate, clean: ohne Duplikate
 
    if DEBUG_INFO: print(f"x_axis_clean ('{kunde_produkt}'):\n{x_axis_clean}")                                                                        # Ausgabe passiert nicht da DEBUG_INFO false ist
 
# Neues und leeres Dictionary wird inalisiert / Name: y_axis / geschweifte Klammer für Dictonary
    y_axis = {}                                                                                                                                       # Der Nutzer füllt das Dictionary aus
 
# Die Werte des Dictionary werden gleich null gesetzt                                                                                                 => war Vorgeben, unnötige Berreinigung/Nulsetzung von den Produkten oder Kunden, da vorher ein leeres Dictonary erstellt wurde
    for x_axis_value in x_axis_clean:       
        y_axis[x_axis_value] = 0            
 
 
# Summe der ausgewählten Variable (Kunde oder Produkt) wird für einzelne Quartale oder für das ganze Jahr gebildet          
    for x_axis_value in x_axis_clean:   # mithilfe der for-Schleife
 
# Besonderheit, da man 'all' nicht als Variable benutzen kann/ Null wird überschrieben
        if quartal == 'all':                                                                                                                          # Wenn Nutzer all/a/al eingibt werden alle quartaler zusammen gerechnet und zugewiesen/ Ablauf wie bei else (nur das alle Quartaler zusammengerechnet werden)
            y_axis[x_axis_value] = df.loc[df[kunde_produkt] == x_axis_value, 'Qrtl 1'].sum() + \
            df.loc[df[kunde_produkt] == x_axis_value, 'Qrtl 2'].sum() + \
            df.loc[df[kunde_produkt] == x_axis_value, 'Qrtl 3'].sum() + \
            df.loc[df[kunde_produkt] == x_axis_value, 'Qrtl 4'].sum()
        else:
            y_axis[x_axis_value] = df.loc[df[kunde_produkt] == x_axis_value, quartal].sum()                                                           # An der Stelle x_axis_value wird in dem dictionary ein Wert zugewesen, nach dem Istgleich wird geschaut wie oft x_axis_value in df vorkommt, anschließend werden diese Werte übernommen und anschließend summiert
 
# Ausgabe der ausgewählten Variable mit den passenden Summe im jeweiligen ausgewählten Quartal
    print(f"\nSumme der Umsätze über das Quartal '{quartal}' für '{kunde_produkt}' ...\n{y_axis}")
 
# Gesamtumsatz der Datei
    gesamtumsatz = df['Qrtl 1'].sum() + df['Qrtl 2'].sum() + df['Qrtl 3'].sum() + df['Qrtl 4'].sum()                                                  # Gesamtumsatz Berechnung mit den Summen der einzelnen Quartale
 
# Berechnung der Summe des ausgewählten Quartals, alle Werte des Quartals aus dem diesem Dictionary 'y_axis' werden addiert und als Gesamtsumme in der variable 'Summe_Quartal' gespeichert
    Summe_Quartal = 0                                                                                                                                 # Variable mit dem Namen 'Summe_Quaratl' initialisiert und auf den Wert 0 gesetzt
    for wert in y_axis.values():                                                                                                                      # Aus dem 'y_axis' Dictionary werden Werte ausgelesen und im nächtes Schritt/Zeile addiert bzw. auf die schon ausgelesenen Werte darauf gerechnet
        Summe_Quartal += wert                                                                                                                         # Aktueller ausgewählter Wert wird zur 'Summe_Quartal' addiert
 
# Berechnung des prozentualen Anteils jedes Kunden oder Produkts am Gesamtumsatz
    prozentualer_anteil= (Summe_Quartal/gesamtumsatz)*100
 
#Ausagbe des prozentualen Anteils des Quartals in Vergleich zum Gesamtumsatz
    print(f"\nProzentualer Anteil des Quartals am Gesamtumsatzes:\n{prozentualer_anteil}")
 
# Optionale Aufgabe 4 (prozentuale Angabe)
     
# Schleife um alle Werte von y_axis zur Verrechnung mit dem ausgewählten Quartal um für jeden Kunden / Produkt einen Prozentsatz auzugegen
    for x_axis_value in y_axis.values():
# Besonderheit bei all, deswegen extra Fall / trettet nur ein wenn 'all' o.a. vom Nutzer eingetragen wird
        if quartal == 'all':
# Prozentualer Anteil wird berechnet im Vergleich zum Gesamtzumsatz
            prozent = (x_axis_value / gesamtumsatz) * 100  
# Der jeweilige Schlüssel wird dazu ermittelt, alos der passende erste Wert im dictionary 
            gesuchter_schluessel = list(y_axis.keys())[list(y_axis.values()).index(x_axis_value)]                                                    # 'list(y_axis.keys())': erzeugt eine Liste aller Schlüssel im Dictionary 'y_axis' / 'list(y_axis.values()).index(x_axis_value)': Gibt die Position des Wertes 'x_axis_value' in der Liste der Werte von 'y_axis' zurück/ 'keys': Extrahiert den Schlüssel an der gefundenen Indexposition
# Ausgabe an den Nutzer
            print(f"\nProzentualer Anteil der Umsätze am Gesamtumsatz über das Quartal '{quartal}' für '{gesuchter_schluessel}' ...\n{prozent}")
        else:
# Prozentualer Anteil wird berechnet vom jeweiligen ausgewählten Quartal
            prozent = (x_axis_value / df[quartal].sum()) * 100  
# Der jeweilige Schlüssel wird dazu ermittelt
            gesuchter_schluessel = list(y_axis.keys())[list(y_axis.values()).index(x_axis_value)]                                                    # Gleiches Spiel wie im if Pfad, nur mit den jeweils ausgewählten einzelnen Quartalen  
# Ausgabe an den Nuter
            print(f"\nProzentualer Anteil der Umsätze am Gesamtumsatz über das Quartal '{quartal}' für '{gesuchter_schluessel}' ...\n{prozent}")
   
    
# Listen bennen
    Personen = list(y_axis.keys()) #Liste aus den Personen aus den Schlüsselwerten (x-Achse)
    Ergebnisse = list (y_axis.values()) # Ergebnisse aus den Werten zu den Schlüsseln (y-Achse)
 
# Punktediagramm aus den ausgewählten Daten Anzeigen lassen
    plt.scatter(Personen, Ergebnisse, color='red', marker='o')                                                                                       # Aufbau des Diagramms: Personen (X-Achse), Ergebnise (Y-Achse), Farbe (rot) marker (In welcher Form die Schnittpunkte angegeben werden sollen)
   
   
    plt.title (f'Umsätze nach {kunde_produkt} im Quartal {quartal}') #Titel des Diagramms
 
# Beschriftung der x-Achsen
    plt.xlabel(kunde_produkt)                                                                                                                        # Beschriftung der x-Achse
    plt.xticks(rotation='vertical')                                                                                                                  # 'plt.sticks': Einstellung der Beschriftung der x-Achse / 'rotation='vertial': Drehung der Beschriftung der x-Achse   
    plt.subplots_adjust(bottom=0.2)                                                                                                                  # 'plt.subplots_adjust': Anpassung verschiedener Parameter / 'bottom=0,4': untere Grenze des Subplots (Gitter oder uch Raster genannt) verschoben
    
# Beschriftung der y-Achse     
    plt.ylabel(quartal)                                                                                                                              # Beschriftung der Y-Achse
 
    plt.grid(True)                                                                                                                                   # Gitterlinien für eine besser Übersicht erzuegen
    plt.show()                                                                                                                                       # Anzeige des Diagramms
 
# Säulendiagramm
    plt.bar(Personen, Ergebnisse, color='blue', alpha=1)                                                                                             # Aufbau des Diagramms: Personen (X-Achse), Ergebnise (Y-Achse), Farbe (blue) alpha (Tranzparenz der Säulen)
    plt.title(f'Umsätze nach {kunde_produkt} im Quartal {quartal}') #Titel des Diagramms
 
# Beschrfitung der x-Achsen
    plt.xlabel(kunde_produkt)                                                                                                                        # x-Achse
    plt.xticks(rotation='vertical')                                                                                                                  # 'plt.sticks': Einstellung der Beschriftung der x-Achse / 'rotation='vertial': Drehung der Beschriftung der x-Achse
    plt.subplots_adjust(bottom=0.4)                                                                                                                  # 'plt.subplots_adjust': Anpassung verschiedener Parameter / 'bottom=0,4': untere Grenze des Subplots (Gitter oder uch Raster genannt) verschoben
    
# Bescriftung y-Achse
    plt.ylabel(quartal)                                                                                                                              
 
    plt.grid(axis='y')                                                                                                                               # auf wechler Achse die Säulen beginnen sollen (senkrecht oder wagegerecht)
    plt.show()                                                                                                                                       # Anzeige des Diagramms
 
    return
 
 
# MAIN
'''plt.ioff()''' #für IOS Nutzer (steht für 'intereactive off': der interaktive Modus von Matplotlib wird deaktiviert, erst bei Anweisung z.B. plt.show() führt dieser sie aus)
if __name__ == "__main__":                                                                                                                            # Bildung der Main
 
# Variablen werden definiert und lokalem Pfad definiert
    wdir = r'C:\GIT\sem1pyhton1'                                                                                      # der eigene Pfad (Hinterlegung auf dem eigenenm Speicher) von diesem Projekt
    infile = r'/101_Umsatzbericht.xlsx'                                                                                                               # Name der Ecxel Datei
    table_name = 'Quelldaten'                                                                                                                         # Bennenung/Beschriftung der Tabelle
 
# Aufrufung der ersten Funktion: try-except Anweisung (Intifizierung von Fehlern im der ersten Funktion)
    try:                                                                                                                                              # Probiert die erste Funktion aufzurufen und durchzuführen
        df, header_row, number_of_headers = xfile_read(wdir, infile, table_name)                                                                                               
    except:                                                                                                                                                                    
        print("Sorry. Could not parse input data correctly. Did you complete your work inside xfile_read()?")                                         # Fehlermeldung für den Benutzer
        raise SystemExit                                                                                                                              # Vorgang wird abgebrochen / raise Anweisung: bestimmte Ausnahmen zu erzwingen
 
# die Variable wird auf False gesetzt, damit die while-Schleife ausgeführt wird (bei True würde die while-Schleife nicht mehr ausgeführt werden)
    parsing_completed = False
 
# while-Schleife erfüllt, da es in der vorherigen Zeile als False identifiziert wurde
    while not parsing_completed:
        parsing_completed = True
 
        print("Kopfzeile: ", end="")                                                                                                                  # Ausgabe: 'Kopfzeile: Spaltennamen' (alles soll in einer Zeile ausgegeben werden)
        print(header_row)
        print("Was möchten Sie analysieren? Kunde (K) oder Produkt (P)?")                                                                             # Frage an den Nutzer
        kunde_produkt = input(">>> ")                                                                                                                 # Nutzer kann seine Antwort eingeben  
 
        print("Welches Quartal möchten Sie betrachten? (1, 2, 3, 4, alle)?")                                                                          # Frage an den Nutzer
        quartal = input(">>> ")                                                                                                                       # Nutzer kann seine Antwort eingeben  
 
# Überprüfung der Eingabe des Nutzers (Groß- und Kleinschreibung ist irrelevant)
        if kunde_produkt == "K" or kunde_produkt == "k":
            kunde_produkt = "Kunde"
        elif kunde_produkt == "P" or kunde_produkt == "p":
            kunde_produkt = "Produkt"
        elif kunde_produkt == "":
            break                                                                                                                                     # Bei keiner Eingabe wird die while-Schleife beendet (break beendet nur die while-Schleife & fragt nicht noch einmal nach, weil die Bedigung False nicht erfüllt ist)
        else:
            parsing_completed = False                                                                                                                 # Bei falscher Eingabe wird erneut gefragt
 
# Überprüfung der Eingabe des Nutzers für das Quartal
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
 
# Bei erfolgreicher Eingabe können wir in die zweite Funktion springen
        if parsing_completed:                                                                                                                          # Variabel (ob True oder False) reagiert /
            plot_client_data(df, kunde_produkt, quartal)                                                                                               # Bei True wird die zweite Funktion wird aufgerufen und die Varibalen werden dieser Funktion übergeben   
        else:
            print("Sorry, no such data to be analyzed. Exiting.")                                                                                      # Bei False wird dieses ausgeben
