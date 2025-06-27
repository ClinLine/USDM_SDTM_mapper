import openpyxl
import pandas as pd
import TS
import TI
import TE

# Define the source json file you like to use
JsonInput = "TestJson/Alexion_NCT04573309_Wilsons.json"
# define the mapping input file
MapInput = "Maps/sdtm_mapping_paths.xlsx"
# Define the resulting output file
Output = "Output/SDTM_Results.xlsx"

if __name__ == "__main__":
    # Create the TS sheet based on the mapping and the json input
    wb = openpyxl.load_workbook(MapInput)
    TS.Create_TS(wb, JsonInput)
    TI.Create_TI(wb, JsonInput)
    TE.Create_TE(wb, JsonInput)
    wb.save(Output)
    wb.close()
