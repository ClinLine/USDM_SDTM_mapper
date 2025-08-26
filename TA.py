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
        elif i == 6: # varname == "TAETORD":
            TaetordCodeSnip = ta_sheet.cell(row=i, column=7).value
            TaetordColumn = j
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
        taetord=definition.Parse_jsonata(codeSnip=TaetordCodeSnip,data=data)        
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
        taetord2 = []
        taetordArm = []
        definition.string_to_nested_list(taetord, taetordArm, taetord2)
        definition.string_to_nested_list(epoch, epochArm, epoch2)  # convert the string to a nested list
        definition.string_to_nested_list(resultEl, resultElArm1, resultEl2)  # convert the string    
        definition.string_to_nested_list(resultElName, resultElArm2, resultElName2)  # convert the string
        ArmNameId2 = {}
        ArmId2 = {}
        Elid2 = {}
        ElNameId2 = {}
        epochId2 = {}
        taetordOrderNext = {}
        taetordOrderNow = {}
        taetordArm3 = {}
        taetordArm4 = []
        taetordId2 = []
        for i in range(len(taetord2)):
            taetordEpoch, taetordNext = definition.get_ID(taetord2[i])  # extracting the ID from the string
            taetordArm2, taetordId = definition.get_ID(taetordArm[i])  # extracting the ID from the string   
            taetordOrderNext[taetordId] = taetordNext  # store the next order in a dictionary with the epoch as key
            taetordOrderNow[taetordId] = taetordEpoch  # store the current order in a dictionary with the epoch as key
            taetordArm3[taetordId] = taetordArm2  # store the arm code in a dictionary with the ID as key
            taetordId2.append(taetordId)  # store the ID in a list
            if taetordArm2 not in taetordArm4:  # check if the arm code is already in the list
                taetordArm4.append(taetordArm2)
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
        taetordId3 = []
        taetordId4 = {}
        for m in taetordArm4:
            currentEpoch = None
            for n in range(len(taetordArm3)):
                for i in range(len(taetordArm3)):
                    if currentEpoch is None:
                        if taetordOrderNext[taetordId2[i]] == "None":
                            if taetordArm3[taetordId2[i]] == m:
                                taetordId3.insert(0,taetordId2[i])
                                currentEpoch = taetordOrderNow[taetordId2[i]]
                                taetordId4[taetordId2[i]] = n + 1
                                break
                    else:
                        if taetordOrderNext[taetordId2[i]] == currentEpoch:
                            if taetordArm3[taetordId2[i]] == m:
                                taetordId3.insert(0,taetordId2[i])
                                currentEpoch = taetordOrderNow[taetordId2[i]]
                                taetordId4[taetordId2[i]] = n + 1
                                break

        values = [v for v in taetordId4.values()]
        val = max(values) + 1

        # Flip the values
        taetordId5 = {
            k: (val - v)
            for k, v in taetordId4.items()
        }
        
        n = 0
        for item in taetordId3:
            x += 1
            ta_sheet.cell(row=x, column=TaetordColumn).value = taetordId5[item]
            ta_sheet.cell(row=x, column=StudyIDColumn).value = studyId
            ta_sheet.cell(row=x, column=DomainColumn).value = DomainResult 
            ta_sheet.cell(row=x, column=ArmNameColumn).value = ArmNameId2[taetordArm3[item]]
            ta_sheet.cell(row=x, column=ArmDescriptionColumn).value = ArmId2[taetordArm3[item]]
            ta_sheet.cell(row=x, column=EtcdColumn).value = ElNameId2[item] 
            ta_sheet.cell(row=x, column=ElementColumn).value = Elid2[item]
            ta_sheet.cell(row=x, column=EpochColumn).value = epochId2[item]
            n += 1
        for j in rem:
            for i in range(2, x + 1):
                ta_sheet.cell(row=i, column=j).value = ""
