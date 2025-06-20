import jsonata
import json
import openpyxl
import pandas as pd
import TS
import TI


# Define the source json file you like to use
JsonInput = "TestJson/ReCoPad.json"
# define the mapping input file
MapInput = "Maps/sdtm_mapping_paths.xlsx"
# Define the resulting output file
Output = "TestJson/SDTM_Results.xlsx"

if __name__ == "__main__":
    # Create the TS sheet based on the mapping and the json input
    wb = openpyxl.load_workbook(MapInput)
    #TS.Create_TS(wb, JsonInput)
    TI.Create_TI(wb, JsonInput)
    wb.save(Output)
