import jsonata
import json
import openpyxl
import pandas as pd

# Define the source json file you like to use
JsonInput = "TestJson/ReCoPad.json"

wb = openpyxl.load_workbook("Maps/sdtm_mapping_paths.xlsx")

# Access the 'TS Parameters' sheet
ts_sheet = wb['TS Parameters']

# Print the value in the first and seventh column of each row in the 'TS Parameters' sheet
with open(JsonInput, 'r') as file:
    data=json.load (file)
    ts_sheet.cell(row=1, column=7).value = "Mapping Results"
    for i in range(2, ts_sheet.max_row + 1):
        MapName = ts_sheet.cell(row=i, column=1).value
        codeSnip = ts_sheet.cell(row=i, column=7).value
        if codeSnip is None:
            result=" "
        else:
            try:
                expr = jsonata.Jsonata(codeSnip)
                result = expr.evaluate(data)            
            except:
                result = "Error in expression for "+ MapName + ": " + codeSnip
        if result is None: result = " "
        print(result)
        result= str(result)
        try:
            result2 = result.replace("’", " ")
        except:
            result2 = None
        if result2 is None: result2= " "
        if result2 is not None:
            ts_sheet.cell(row=i, column=7).value = result2            
         #   print(result2)
wb.save("Output/sdtm_mapping_results.xlsx")