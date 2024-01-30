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

    if DEBUG_INFO: print(f"df: {df}")

    print(list(in_file))
    
    # if DEBUG_INFO: print(f"df: {df}")

    # printing the header row as list of values, e.g. 
    # header_row:  ['Produkt', 'Kunde', 'Qrtl 1', 'Qrtl 2', 'Qrtl 3', 'Qrtl 4']
 
    # printing the number of values in the header row, e.g.
    # number_of_headers:  6
 
    # if DEBUG_INFO: print(f"header_row: {header_row} \nnumber_of_headers: {number_of_headers}")

    return df, header_row, number_of_headers

# MAIN
if __name__ == "__main__":

    # Initialize working directory, file and tab name
    wdir = r'D:\Lehre\Sem 1\Python' 
    infile = r'\101_Umsatzbericht.xltx'
    table_name = 'Quelldaten'
    xfile_read(wdir, in_file, tabble_name)
    #try:
     #   df, header_row, number_of_headers = xfile_read(wdir, infile, table_name)
    #except:
        #print("Sorry. Could not parse input data correctly. Did you complete your work inside xfile_read()?")
        #raise SystemExit
