import openpyxl
import pandas as pd
import TS
import TI
import TE
import TA
import TV
from create_define import Create_Define

# Define the source json file you like to use
JsonInput = "TestJson/CDISC_Pilot_Study_v4_FIXED_adapted.json"
# define the mapping input file
MapInput = "Maps/sdtm_mapping_paths.xlsx"
# Define the resulting output file
Output = "Output/SDTM_Results.xlsx"

if __name__ == "__main__":
    # Create the TS sheet based on the mapping and the json input
    code_lists_map={}
    wb = openpyxl.load_workbook(MapInput)
    TS.Create_TS(wb, JsonInput)
    ti_var = TI.Create_TI(wb, JsonInput)
    TE.Create_TE(wb, JsonInput)
    ta_var, ta_codes = TA.Create_TA(wb, JsonInput)
    TV.Create_TV(wb, JsonInput)    
    code_lists_map.update(ta_codes)
    Create_Define(wb,ta_var,ti_var,code_lists_map)
    wb.save(Output)
    wb.close()
