import jsonata
import json
import openpyxl
import pandas as pd

# Define the source json file you like to use
JsonInput = "TestJson/ReCoPad.json"

wb = openpyxl.load_workbook("Maps/sdtm_mapping_paths.xlsx")

# Access the 'TS Parameters' sheet
ti_sheet = wb['TI']
for i in range(2, ti_sheet.max_row + 1):
    j=i-1
    #swap the rows and columns in the TI sheet
    varName=ti_sheet.cell(row=i, column=1).value
    ti_sheet.cell(row=1, column=j).value = varName    
    if varName == "STUDYID":
        StudyIdCodeSnip = ti_sheet.cell(row=i, column=7).value
        StudyIDColumn = j
    if varName == "DOMAIN":
        DomainResult =  ti_sheet.cell(row=i, column=8).value
        DomainColumn = j
    if varName == "TIVERS":
        VersionResult =  ti_sheet.cell(row=i, column=8).value
        VersionColumn = j

def string_to_list(input, result):
    n = 0 #letter it is looking at
    while input[n] != "}": #looking for the end of the list
        if input[n-1:n+3] == "', '" or input[n] in ["{"]: # looking for the start of a new item in the list
            n += 1
            m = n
            while m+1 < len(input) and input[m+1] not in ("}") and input[m:m+3] not in ("', '"): # looking for the end of the item
                m += 1
            result.append(input[n:m]) # appending the item to the list
            n = m + 1
            #tst=input[n-1:n+3] 
            #print("tst: ", tst)
        else: 
            n += 1
            

def Parse_jsonata(codeSnip):
    if codeSnip is None:
        result = " "
    else:
        try:
            expr = jsonata.Jsonata(codeSnip)
            result = expr.evaluate(data)  
        except:
            result = "Error in expression for "+ varName + ": " + codeSnip
    if result is None: result = " "
    result= str(result)
    if result == "": result = " "
    if result == "{}": result = " "
    try:
        result0 = result.replace("’", " ")
    except:
        result0 = ""
    if result0 == "": result0= " "
    if result0[0] == "[": 
            result0 = result0[1:-1]
            result0 = result0.replace("}, {", ", ")
    return result0
    
# Print the value in the first and seventh column of each row in the 'TS Parameters' sheet
with open(JsonInput, 'r') as file:
    data=json.load (file)
    studyId=Parse_jsonata(codeSnip=StudyIdCodeSnip) 
    
    print("StudyIdCodeSnip: ", StudyIdCodeSnip)
    print("StudyIdColumn: ", StudyIDColumn)
    print("StudyId: ", studyId)
    Version=Parse_jsonata(VersionResult)
    for i in range(2, ti_sheet.max_row + 1):
        # Get all the mapping information from the TS Parameters sheet
        if i not in [StudyIDColumn+1, DomainColumn+1, VersionColumn+1]:
            codeSnip = ti_sheet.cell(row=i, column=7).value
            result2=Parse_jsonata(codeSnip=codeSnip)
            x=1
            if result2 != " ":
                if result2[0] == "{":  # check if the result is a list
                    result3 = []
                    string_to_list(result2, result3)  # convert the string to a list
                    #if i==5: 
                       # print ("result2: ", result2)
                       # print ("result3: ", result3)
                    # filling ts sheet if it is a list 
                    for j in range(0, len(result3)):
                        x += 1
                        c = i-1
                        ti_sheet.cell(row=x, column=c).value = result3[j]
                        ti_sheet.cell(row=x, column=StudyIDColumn).value = studyId
                        ti_sheet.cell(row=x, column=DomainColumn).value = DomainResult
                        ti_sheet.cell(row=x, column=VersionColumn).value = Version
            else:
                # filling ts sheet if it is not a list
                x += 1
                ti_sheet.cell(row=x, column=c).value = result2
                ti_sheet.cell(row=x, column=StudyIDColumn).value = studyId
                ti_sheet.cell(row=x, column=DomainColumn).value = DomainResult
                ti_sheet.cell(row=x, column=VersionColumn).value = Version
    file.close
wb.save("Output/sdtm_mapping_results.xlsx")