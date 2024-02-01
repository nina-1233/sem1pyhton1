# Further readings: https://www.datacamp.com/community/tutorials/python-excel-tutorial
# WICHTIG: Neben pandas und matplotlib muss zusätzlich noch die openpyxl Bibliothek 
# imstalliert werden. Diese wird im Hintergrund für das Öffnen der Exceldatei in der
# Anweisung pd.ExcelFile(in_file) in xfile_read() benötigt.

import pandas as pd
import matplotlib.pyplot as plt

DEBUG_INFO = True

'''
xfile_read: reading and parsing an input excel file
input parameters: 
- wdir: working directory
- in_file: input file name  
- tabble_name: tab name of excel sheet to be analysed (parsed)
output params:
- df: data frame (for further analysis of Excel data in concern)
- header_row: list of header entries of sheet in concern
- number_of_headers: number of entries in header_row list
'''
def xfile_read(wdir, in_file, tabble_name):

    in_file = wdir + in_file
    print(f"Loading file '{in_file}' ...")

    xl = pd.ExcelFile(in_file)
    #if DEBUG_INFO: print("All sheet names: ", xl.sheet_names)

    if tabble_name == "": tabble_name = 'Quelldaten'
    print(f"Parsing data from sheet '{tabble_name}'")
    df = pd.read_excel(in_file)

    header_row = df.columns.tolist()
    number_of_headers= len(header_row)
    print(header_row)
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

    if kunde_produkt == 'Kunde':
        x_axis_raw = df['Kunde']
    if kunde_produkt == 'Produkt':
        x_axis_raw = df['Produkt']
    
    Liste = list(x_axis_raw)
    x_axis_clean = set(Liste)
 
    #if DEBUG_INFO: print(f"x_axis_clean ('{kunde_produkt}'):\n{x_axis_clean}")

    y_axis = {}
    for x_axis_value in x_axis_clean:
        y_axis[x_axis_value] = 0
    
    if kunde_produkt == 'Kunde':
         for x_axis_value in x_axis_clean:
            if quartal == 'Qrtl 1':
                y_axis[x_axis_value] = df.loc[df['Kunde'] == x_axis_value, 'Qrtl 1'].sum()
            if quartal == 'Qrtl 2':
                y_axis[x_axis_value] = df.loc[df['Kunde'] == x_axis_value, 'Qrtl 2'].sum()
            if quartal == 'Qrtl 3':
                y_axis[x_axis_value] = df.loc[df['Kunde'] == x_axis_value, 'Qrtl 3'].sum()    
            if quartal == 'Qrtl 4':
                y_axis[x_axis_value] = df.loc[df['Kunde'] == x_axis_value, 'Qrtl 4'].sum() 
            if quartal == 'all':    
                y_axis[x_axis_value] = df.loc[df['Kunde'] == x_axis_value, 'Qrtl 1'].sum() + df.loc[df['Kunde'] == x_axis_value, 'Qrtl 2'].sum() + df.loc[df['Kunde'] == x_axis_value, 'Qrtl 3'].sum() +df.loc[df['Kunde'] == x_axis_value, 'Qrtl 4'].sum() 

    
    if kunde_produkt == 'Produkt':
         for x_axis_value in x_axis_clean:
            if quartal == 'Qrtl 1':
                y_axis[x_axis_value] = df.loc[df['Produkt'] == x_axis_value, 'Qrtl 1'].sum()
            if quartal == 'Qrtl 2':
                y_axis[x_axis_value] = df.loc[df['Produkt'] == x_axis_value, 'Qrtl 2'].sum()
            if quartal == 'Qrtl 3':
                y_axis[x_axis_value] = df.loc[df['Produkt'] == x_axis_value, 'Qrtl 3'].sum()    
            if quartal == 'Qrtl 4':
                y_axis[x_axis_value] = df.loc[df['Produkt'] == x_axis_value, 'Qrtl 4'].sum() 
            if quartal == 'all':
                y_axis[x_axis_value] = df.loc[df['Produkt'] == x_axis_value, 'Qrtl 1'].sum() + df.loc[df['Produkt'] == x_axis_value, 'Qrtl 2'].sum() +  df.loc[df['Produkt'] == x_axis_value, 'Qrtl 3'].sum() + df.loc[df['Produkt'] == x_axis_value, 'Qrtl 4'].sum()
    
    print(f"\nSumme der Umsätze über das Quartal '{quartal}' für '{kunde_produkt}' ...\n{y_axis}")

    # Now, we are ready to plot y_axis (as value dictionary). We use the dictionary keys as x-values,
    # dictionary values as y-values and function plt.plot(), i.e. matplotlib.pyplot.plot().
    # https://www.w3schools.com/python/python_dictionaries.asp
    # https://matplotlib.org/stable/api/_as_gen/matplotlib.pyplot.plot.html 
 
    # Ihr Code folgt hier: ...

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
