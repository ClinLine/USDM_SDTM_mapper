import jsonata
import json
import openpyxl
import pandas as pd
import shutil

# Define the source json file you like to use
JsonInput = "TestJson/ReCoPad.json"

source_file = "Maps/sdtm_mapping_paths.xlsx"
destination_file = "Output/sdtm_mapping_results.xlsx"

# shutil.copy(source_file, destination_file)

wb = openpyxl.load_workbook("Maps/sdtm_mapping_paths.xlsx")

# Access the 'TS Parameters' sheet
ts_sheet = wb['TS Parameters']

# Print the value in the first and seventh column of each row in the 'TS Parameters' sheet
with open(JsonInput, 'r') as file:
    data=json.load (file)
    for i in range(2, ts_sheet.max_row + 1):
        codeSnip = ts_sheet.cell(row=i, column=7).value
        result2=" "
        try:
            expr = jsonata.Jsonata(codeSnip)
            result = expr.evaluate(data)            
        except:
            result = None
        try:
            result2 = result.replace("’", " ")
        except:
            result2 = None
        if result is not None:
            ts_sheet.cell(row=i, column=8).value = result2
            print(result2)
wb.save("Output/sdtm_mapping_results.xlsx")