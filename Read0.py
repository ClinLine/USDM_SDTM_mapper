import jsonata
import json
import openpyxl
import pandas as pd

wb = openpyxl.load_workbook("Maps/sdtm_mapping_paths.xlsx")

# Access the 'TS Parameters' sheet
ts_sheet = wb['TS Parameters']

# Print the value in the first and seventh column of each row in the 'TS Parameters' sheet
with open('ReCoPad.json', 'r') as file:
    data=json.load (file)
    for i in range(2, ts_sheet.max_row + 1):
        codeSnip = ts_sheet.cell(row=i, column=7).value
        try:
            expr = jsonata.Jsonata(codeSnip)
            result = expr.evaluate(data)
            result2 = result.replace("’", " ")
        except:
            result2 = "No result"
        if result2 is not None:
            ts_sheet.cell(row=i, column=8).value = result2
            print(result2)
wb.save("Maps/sdtm_mapping_paths.xlsx")