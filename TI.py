import jsonata
import json
import pandas as pd
import definition

def Create_TI(wb, JsonInput):
    ti_sheet = wb['TI']
    StudyIdCodeSnip = ""
    DomainResult = ""
    VersionCodeSnip = ""

    for r in range(2, ti_sheet.max_row+2):
        if ti_sheet.cell(row=r, column=1).value is not None:
            overflowcolumn = r

    for i in range(2, overflowcolumn + 1):
        j=i-1
        #swap the rows and columns in the TI sheet
        varName=ti_sheet.cell(row=i, column=1).value
        ti_sheet.cell(row=1, column=j).value = varName    

        # Get standard row value information for TI
        if varName == "STUDYID":
            StudyIdCodeSnip = ti_sheet.cell(row=i, column=7).value
            StudyIDColumn = j
        if varName == "DOMAIN":
            DomainResult =  ti_sheet.cell(row=i, column=8).value
            DomainColumn = j
        if varName == "TIVERS":
            VersionCodeSnip =  ti_sheet.cell(row=i, column=7).value
            VersionColumn = j
        # Get variable row value information for TI
        if varName == "IETESTCD":
            TestCdColumn = j
        if varName == "IETEST":
            TestColumn = j
        if varName == "IECAT":
            CatColumn = j
        if varName == "IESCAT":
            ScatColumn = j
        if varName == "TIRL":
            RlColumn = j

    # Print the value in the first and seventh column of each row in the 'TS Parameters' sheet
    with open(JsonInput, 'r') as file:
        data=json.load (file)
        # Get all the standard mapping information from the TI sheet
        studyId=definition.Parse_jsonata(codeSnip=StudyIdCodeSnip,data=data)         
       # print ("V:" + VersionCodeSnip)
        versionResult=definition.Parse_jsonata(codeSnip=VersionCodeSnip,data=data)
       # print("TI StudyId: ", studyId)
        
        TestCdResult=definition.Parse_jsonata(codeSnip=ti_sheet.cell(row=TestCdColumn+1, column=7).value,data=data)
        TestResult=definition.Parse_jsonata(codeSnip=ti_sheet.cell(row=TestColumn+1, column=7).value,data=data)
        CatResult=definition.Parse_jsonata(codeSnip=ti_sheet.cell(row=CatColumn+1, column=7).value,data=data)
        ScatResult=definition.Parse_jsonata(codeSnip=ti_sheet.cell(row=ScatColumn+1, column=7).value,data=data)
        RlResult=definition.Parse_jsonata(codeSnip=ti_sheet.cell(row=RlColumn+1, column=7).value,data=data)

        #transform the results into ID based lists
        TestCdList = {}
        definition.string_to_ID_list(TestCdResult, TestCdList)  # convert the string to a list
        TestResultList = {}
        definition.string_to_ID_list(TestResult, TestResultList)  # convert the string to a
        CatResultList = {}
        definition.string_to_ID_list(CatResult, CatResultList)  # convert the string to a
        ScatResultList = {}
        definition.string_to_ID_list(ScatResult, ScatResultList)  # convert the string to a
        RlResultList = {}
        definition.string_to_ID_list(RlResult, RlResultList)  # convert the string to a
        skip=0

        #Resolve tags in the restResultList
        for i in TestResultList:
            result4=TestResultList[i]
            result4=definition.ResolveTag(result4,data)
            TestResultList[i]=result4

        #Fill sheet
        x=2
        for i in TestCdList:
            ti_sheet.cell(row=x, column=StudyIDColumn).value = studyId
            ti_sheet.cell(row=x, column=DomainColumn).value = DomainResult
            ti_sheet.cell(row=x, column=VersionColumn).value = versionResult    
            ti_sheet.cell(row=x, column=TestCdColumn).value = TestCdList[i]
            # Break result in strings of max 200 characters and fill in the TI sheet
            breakoff = 0
            if i in TestResultList: 
                result4=TestResultList[i]
                while len(result4) > 0:
                    k=len(result4)
                    if k>200:
                        for k in range(200, 0, -1):
                            if result4[k] == " ": 
                                printResult4 = result4[0:k]
                                break                        
                    else:
                        printResult4 = result4
                    if breakoff == 0:
                        ti_sheet.cell(row=x, column=TestColumn).value = printResult4
                    else:
                        #print(str(breakoff) + ":" + printResult4)    
                        ti_sheet.cell(row=x, column=overflowcolumn-1+breakoff).value = printResult4
                    result4 = result4[len(printResult4):]
                    if len(result4) <4: break
                    breakoff += 1
               
            if i in CatResultList: 
                ti_sheet.cell(row=x, column=CatColumn).value = CatResultList[i]
            else:
                ti_sheet.cell(row=x, column=CatColumn).value = " "
            if i in ScatResultList: 
                ti_sheet.cell(row=x, column=ScatColumn).value = ScatResultList[i]
            else:
                ti_sheet.cell(row=x, column=ScatColumn).value = " "
            if i in RlResultList: 
                ti_sheet.cell(row=x, column=RlColumn).value = RlResultList[i]  
            else:
                ti_sheet.cell(row=x, column=RlColumn).value = " "
            
            x += 1
    
        # Rename overflow columns
        if ti_sheet.max_column > 8:   
            for i in range(9, ti_sheet.max_column):
                newname = f"IETEST{i-7}"
                ti_sheet.cell(row=1, column=i).value = newname