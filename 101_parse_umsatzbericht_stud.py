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

    # Assign spreadsheet filename from wdir to `in_file`
    # Define output excel file accordingly
    in_file = wdir + in_file

    # Load spreadsheet
    print(f"Loading file '{in_file}' ...")
    xl = pd.ExcelFile(in_file)

    # Print all sheet names
    if DEBUG_INFO: print("All sheet names: ", xl.sheet_names)

    # Load a sheet into a DataFrame by name: df
    # Print complete table/DataFrame
    if tabble_name == "": tabble_name = 'Quelldaten'
    print(f"Parsing data from sheet '{tabble_name}'")
    # Generating a data frame df out of tabble_name using xl.parse()
    # https://pandas.pydata.org/pandas-docs/stable/reference/api/pandas.DataFrame.html

    # Ihr Code folgt hier: ...

    df.x1.parsa(tabble_name)
    if DEBUG_INFO: print(f"df: {df}")

    number_of_headers= len(header_row)
    print(number_of_headers)
    
    if DEBUG_INFO: print(f"df: {df}")

    # printing the header row as list of values, e.g. 
    # header_row:  ['Produkt', 'Kunde', 'Qrtl 1', 'Qrtl 2', 'Qrtl 3', 'Qrtl 4']
    header_row = list(in_file)
    print(header_row)
 
    # printing the number of values in the header row, e.g.
    # number_of_headers:  6
    number_of_headers= len(header_row)
    print(number_of_headers)
 
    if DEBUG_INFO: print(f"header_row: {header_row} \nnumber_of_headers: {number_of_headers}")

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
    # Plot data for each single kunde_produkt (Kunde oder Produkt)
    # First, we have to extract kunde_produkt (Kunde oder Produkt) entries from df and save them to x_axis_raw
    # https://pandas.pydata.org/pandas-docs/stable/reference/api/pandas.DataFrame.html

    # Ihr Code folgt hier: ...

    # erasing double entries to yield in unique data for each kunde_produkt (Kunde oder Produkt)
    # we will have to use a combination of list() and set() to get the resulting list 'x_axis_clean'

    # Ihr Code folgt hier: ...
 
    # if DEBUG_INFO: print(f"x_axis_clean ('{kunde_produkt}'):\n{x_axis_clean}")

    # initializing values dictionary (y_axis == quartal) for each (cleaned) kunde_produkt (Kunde oder Produkt)
    y_axis = {}
    for x_axis_value in x_axis_clean:
        y_axis[x_axis_value] = 0

    # Summing up quartal values for each kunde_produkt in values dictionary
    # To accomplish this, we have to iterate and sum up over all values in list 'x_axis_raw', thus:
    # for x_axis_value in x_axis_raw:

    # Case 1: quartal == "all": Berechnung von y_axis[x_axis_value], also des y-Wertes am x-Wert 'x_axis_value' (z.B. Kunde 'ANTON')
    # Den zu brechnenden y-Wert müssen wir über alle Quartale ("Qrtl 1" bis "Qrtl 4") und für alle 
    # 'x_axis_value' (also an allen Stellen, an denen z.B. Kunde 'ANTON' erwähnt wird) summieren. 
    # Die Werte selbst beziehen wir über das data frame df
    
    # Case 2: quartal = "Qrtl 1", etc.: Berechnung von y_axis[x_axis_value], also des y-Wertes am x-Wert 'x_axis_value' (z.B. Kunde 'ANTON')
    # Den zu brechnenden y-Wert müssen wir über das Quartal 'quartal' (Parameter!) und für alle 
    # 'x_axis_value' (also an allen Stellen, an denen z.B. Kunde 'ANTON' erwähnt wird) summieren. 
    # Die Werte selbst beziehen wir über das data frame df

    # Ihr Code folgt hier: ...

    # y_axis (value dictionary) for all x-values in x_axis_clean and quartal '{quartal}': {y_axis}
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
    infile = r'\101_Umsatzbericht.xltx'
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
