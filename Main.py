import jsonata
import json
import openpyxl
import pandas as pd
import TS
import TI

# general function(s)
def string_to_list(input, result):
    n = 0 #letter it is looking at
    while input[n] != "}": #looking for the end of the list
        if input[n-2:n+2] == "', '" or input[n] in ["{"]: # looking for the start of a new item in the list
            n += 1
            m = n
            while m+1 < len(input) and input[m+1] not in ("}") and input[m:m+3] not in ("', '"): # looking for the end of the item
                m += 1
            result.append(input[n:m]) # appending the item to the list
            n = m + 1
        else: 
            n += 1

# Define the source json file you like to use
JsonInput = "TestJson/EliLilly_NCT03421379_Diabetes.json"
# define the mapping input file
MapInput = "Maps/sdtm_mapping_paths.xlsx"
# Define the resulting output file
Output = "TestJson/SDTM_Results.xlsx"

if __name__ == "__main__":
    # Create the TS sheet based on the mapping and the json input
    wb = openpyxl.load_workbook(MapInput)
    TS.Create_TS(wb, JsonInput)
    TI.Create_TI(wb, JsonInput)
    wb.save(Output)
