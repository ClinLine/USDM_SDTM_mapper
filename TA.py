import jsonata
import json
import openpyxl
import pandas as pd
import definition

def Create_TA(wb, JsonInput):
    ta_sheet = wb['TA']
    r = ta_sheet.max_row
    rem = []
    for i in range(2, ta_sheet.max_row + 1):
        j=i-1
        #swap the rows and columns in the TI sheet
        varName=ta_sheet.cell(row=i, column=1).value
        ta_sheet.cell(row=1, column=j).value = varName    
        if i == 2: # varName == "STUDYID":
            StudyIdCodeSnip = ta_sheet.cell(row=i, column=7).value
            StudyIDColumn = j
        elif i == 3: # varName == "DOMAIN":
            DomainResult =  ta_sheet.cell(row=i, column=8).value
            DomainColumn = j
        elif i == 4: # varName == "ARMCD":
            ArmNameCodeSnip = ta_sheet.cell(row=i, column=7).value
            ArmNameColumn = j
        elif i == 5: # varName == "ARM":
            ArmDescriptionCodeSnip = ta_sheet.cell(row=i, column=7).value
            ArmDescriptionColumn = j
        elif i == 7: # varName == "ETCD":
            EtcdCodeSnip = ta_sheet.cell(row=i, column=7).value
            EtcdColumn = j
        elif i == 8: # varName == "ELEMENT":
            ElementCodeSnip = ta_sheet.cell(row=i, column=7).value
            ElementColumn = j
        elif i == 11: # varName == "EPOCH":
            EpochCodeSnip = ta_sheet.cell(row=i, column=7).value
            EpochColumn = j
        else:
            rem.append(j)

    for i in range(r, ta_sheet.max_column):
        for j in range(1, r):
            ta_sheet.cell(row=j, column=i).value = ""
    # create empty id array for checking value alignment in different columns
    id = []
    # Print the value in the first and seventh column of each row in the 'TS Parameters' sheet
    with open(JsonInput, 'r') as file:
        data=json.load (file)
        studyId=definition.Parse_jsonata(codeSnip=StudyIdCodeSnip,data=data)         
        x=1
        resultArm=definition.Parse_jsonata(codeSnip=ArmDescriptionCodeSnip,data=data)
        resultArmName=definition.Parse_jsonata(codeSnip=ArmNameCodeSnip,data=data)
        epoch = definition.Parse_jsonata(codeSnip=EpochCodeSnip,data=data)
        resultArm2 = []
        resultArmName2 = []
        definition.string_to_list(resultArm, resultArm2)  # convert the string to a list
        definition.string_to_list(resultArmName, resultArmName2)  # convert the string to a list
        resultEl=definition.Parse_jsonata(codeSnip=EtcdCodeSnip,data=data)
        resultElName=definition.Parse_jsonata(codeSnip=ElementCodeSnip,data=data)
        resultEl2 = []
        resultElName2 = []
        resultElArm1 = []
        epoch2 = []
        epochArm = []
        resultElArm2 = []
        definition.string_to_nested_list(epoch, epochArm, epoch2)  # convert the string to a nested list
        definition.string_to_nested_list(resultEl, resultElArm1, resultEl2)  # convert the string    
        definition.string_to_nested_list(resultElName, resultElArm2, resultElName2)  # convert the string
        ArmNameId2 = {}
        ArmId2 = {}
        Elid2 = {}
        ElNameId2 = {}
        epochId2 = {}
        for item in epoch2:
            epochId, epoch3 = definition.get_ID(item)  # extracting the ID from the string
            epochId2[epochId] = epoch3  # store the epoch code in a dictionary with the ID as key
        for item in resultArmName2:  
            ArmNameId, resultArmName3 = definition.get_ID(item)
            ArmNameId2[ArmNameId] = resultArmName3  # store the arm name in a dictionary with the ID as key
        for item in resultArm2:
            ArmId, resultArm3 = definition.get_ID(item)  # extracting the ID from the string
            ArmId2[ArmId] = resultArm3  # store the arm code in a dictionary with the ID as key
        for item in resultEl2:
            ElId, resultEl3 = definition.get_ID(item)  # extracting the ID from the string
            Elid2[ElId] = resultEl3  # store the element code in a dictionary with the ID as key
        for item in resultElName2:
            ElNameId, resultElName3 = definition.get_ID(item)
            ElNameId2[ElNameId] = resultElName3  # store the element name in a dictionary with the ID as key
        for item in resultElArm1:
            x += 1
            ArmId9, ElId9 = definition.get_ID(item)
            ta_sheet.cell(row=x, column=StudyIDColumn).value = studyId
            ta_sheet.cell(row=x, column=DomainColumn).value = DomainResult 
            ta_sheet.cell(row=x, column=ArmNameColumn).value = ArmNameId2[ArmId9]
            ta_sheet.cell(row=x, column=ArmDescriptionColumn).value = ArmId2[ArmId9]
            ta_sheet.cell(row=x, column=EtcdColumn).value = ElNameId2[ElId9] 
            ta_sheet.cell(row=x, column=ElementColumn).value = Elid2[ElId9]
            ta_sheet.cell(row=x, column=EpochColumn).value = epochId2[ElId9]
        for j in rem:
            for i in range(2, x + 1):
                ta_sheet.cell(row=i, column=j).value = ""
