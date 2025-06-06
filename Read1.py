# TEST jSON CODE EXPERIMENTATION

import jsonata
import json
import openpyxl

# close old json file

# Read JSON file

with open('EliLilly_NCT03421379_Diabetes.json', 'r') as file:
    data=json.load (file)
# expr= jsonata.Jsonata("$count(study.versions.studyDesigns.arms)")
# result=expr.evaluate(data)
# print(result)

wb = openpyxl.load_workbook('Maps/sdtm_mapping_paths.xlsx')

# Start with TS Summary domain
ts_sheet = wb['TS Parameters']

#print(str(ts_sheet['G2'].value))
#code_snip= ts_sheet['G5'].value
#expr= jsonata.Jsonata("$count(study.versions.studyDesigns.arms)")
#result=expr.evaluate(data)
#print(result)


with open('ReCoPad.json', 'r') as file:
    data=json.load (file)
    for row in range(30, 32):
        codeSnip= ts_sheet[f'G{row}'].value
        varDesc= ts_sheet[f'A{row}'].value
        # print(codeSnip)
        try:
            expr = jsonata.Jsonata(codeSnip)
            result=expr.evaluate(data)
            result2 = result.replace("’", " ")
            if result is not None:
                print(result2)
        except:
            result2="Error in expression"
        # if result is not None:
    