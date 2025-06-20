import jsonata
import json
import openpyxl

def get_ID(ID_string):
    if len(ID_string) <2: # if the ID string is None, return empty strings
        return "", ""
    else:
        o = 1 #letter it is looking at
        while ID_string[o] != ":" and o+1 < len(ID_string): #looking for the end of the ID
            o += 1
        if o+1 == len(ID_string): # if the ID is not found, return empty strings
            return "", ID_string
        else:
            Id = ID_string[1:o-1] # extracting the ID from the string
            if ID_string[-2:-1] == "'":
                ID_less = ID_string[o+3:-2]  # extracting the ID without the prefix
            else:
                ID_less = ID_string[o+3:]  # extracting the ID without the prefix
            return Id, ID_less

# general function(s)
def string_to_list(input, result):
    n = 0 #letter it is looking at
    while input[n] != "}": #looking for the end of the list
        if input[n-2:n+2] == "', '" or input[n] in ["{"]: # looking for the start of a new item in the list
            n += 1
            m = n
            while m+1 < len(input) and input[m+1] not in ("}") and input[m:m+3] not in ("', '"): # looking for the end of the item
                m += 1
            result.append(input[n:m]) # appending the item to the list
            n = m + 1
        else: 
            n += 1

def Parse_jsonata(my_sheet,row,column,data):
    codeSnip = my_sheet.cell(row=row, column=column).value
    if codeSnip is None:
        result = " "
    else:
        try:
            expr = jsonata.Jsonata(codeSnip)
            result = expr.evaluate(data)  
        except:
            result = "Error in expression for: " + codeSnip
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
            result0 = result0.replace("}, {", " , ")
    return result0
    
def Create_TS(wb, JsonInput):
    ts0_sheet = wb['TS']
    ts_sheet = wb['TS Parameters']

    for i in range(2, ts0_sheet.max_row + 1):
        j=i-1
        #swap the rows and columns in the TS sheet
        varName=ts0_sheet.cell(row=i, column=1).value
        ts0_sheet.cell(row=1, column=j).value = varName
        DomainResult = " "
        StudyIdCodeSnip = " "
        # Get standard row value information for TS
        if varName == "STUDYID":
            StudyIdCodeSnip = ts0_sheet.cell(row=i, column=7).value
        if varName == "DOMAIN":
            DomainResult =  ts0_sheet.cell(row=i, column=8).value

        with open(JsonInput, 'r') as file:
            data=json.load (file)
            # Get StudyId first and start with a row id for the TS sheet            
            try:
                expr = jsonata.Jsonata (StudyIdCodeSnip)
                studyId= expr.evaluate(data)
            except:
                studyId = "no data "
            x=1
            # Then continue with the mappings in the TS Parameters sheet
            ts_sheet.cell(row=1, column=7).value = "Mapping Results"
            for i in range(2, ts_sheet.max_row + 1):
                # Get all the mapping information from the TS Parameters sheet
                MapName = ts_sheet.cell(row=i, column=1).value
                MapCode = ts_sheet.cell(row=i, column=2).value
                nfValue = ts_sheet.cell(row=i, column=8).value
                result2=Parse_jsonata(ts_sheet,i,7,data)
                resultCd=Parse_jsonata(ts_sheet,i,9,data)
                resultCdRef=Parse_jsonata(ts_sheet,i,10,data)
                resultCdVer=Parse_jsonata(ts_sheet,i,11,data)
                codeSnip = ts_sheet.cell(row=i, column=7).value        
                codeSnipCd = ts_sheet.cell(row=i, column=9).value   
                codeSnipCdRef = ts_sheet.cell(row=i, column=10).value   
                codeSnipCdVer = ts_sheet.cell(row=i, column=11).value   
        
                # replace the apostrophes with spaces
                if result2 != " " or nfValue != " ":
                    if result2[0] == "{":  # check if the result is a list
                        result3 = []
                        string_to_list(result2, result3)  # convert the string to a list
                        if resultCd != " ":
                            resultCd2 = []
                            string_to_list(resultCd, resultCd2)  # convert the string to a list
                            resultCdRef2 = []
                            string_to_list(resultCdRef, resultCdRef2)  # convert the string to a list
                            resultCdVer2 = []
                            string_to_list(resultCdVer, resultCdVer2)  # convert the string to
                        # filling ts sheet if it is a list 
                        for j in range(0, len(result3)):
                            x += 1
                            id, result4 = get_ID(result3[j])  # extract the ID from the result
                            ts0_sheet.cell(row=x, column=1).value = " "
                            ts0_sheet.cell(row=x, column=1).value = studyId
                            ts0_sheet.cell(row=x, column=2).value = DomainResult   
                            ts0_sheet.cell(row=x, column=3).value = j + 1
                            ts0_sheet.cell(row=x, column=4).value = id
                            ts0_sheet.cell(row=x, column=5).value = MapCode    
                            ts0_sheet.cell(row=x, column=6).value = MapName                
                            ts0_sheet.cell(row=x, column=7).value = result4  
                            ts0_sheet.cell(row=x, column=8).value = " "   
                            if result2 == " ": ts0_sheet.cell(row=x, column=8).value = nfValue
                            if resultCd != " ":
                                id, resultcd4 = get_ID(resultCd2[j])
                                id, resultCdRef4 = get_ID(resultCdRef2[j]) 
                                id, resultCdVer4 = get_ID(resultCdVer2[j]) 
                                ts0_sheet.cell(row=x, column=9).value = resultcd4                  
                                ts0_sheet.cell(row=x, column=10).value = resultCdRef4 
                                ts0_sheet.cell(row=x, column=11).value = resultCdVer4
                    else:
                        # filling ts sheet if it is not a list
                        x += 1
                        id, result2 = get_ID(result2)  # extract the ID from the result
                        ts0_sheet.cell(row=x, column=1).value = " "
                        ts0_sheet.cell(row=x, column=1).value = studyId
                        ts0_sheet.cell(row=x, column=2).value = DomainResult   
                        ts0_sheet.cell(row=x, column=3).value = " "
                        ts0_sheet.cell(row=x, column=4).value = " "
                        ts0_sheet.cell(row=x, column=5).value = MapCode    
                        ts0_sheet.cell(row=x, column=6).value = MapName                
                        ts0_sheet.cell(row=x, column=7).value = result2   
                        ts0_sheet.cell(row=x, column=8).value = " "   
                        if result2 == " ": ts0_sheet.cell(row=x, column=8).value = nfValue	
                        if resultCd != " ":
                            id, resultCd = get_ID(resultCd)
                            id, resultCdRef = get_ID(resultCdRef)
                            id, resultCdVer = get_ID(resultCdVer)
                            ts0_sheet.cell(row=x, column=9).value = resultCd                   
                            ts0_sheet.cell(row=x, column=10).value = resultCdRef  
                            ts0_sheet.cell(row=x, column=11).value = resultCdVer
                # filling TS Parameters sheet
                ts_sheet.cell(row=i, column=7).value = result2
                ts_sheet.cell(row=i, column=8).value = " "   
                if result2 == " ": ts_sheet.cell(row=i, column=8).value = nfValue	
                ts_sheet.cell(row=i, column=9).value = resultCd  
                ts_sheet.cell(row=i, column=10).value = resultCdRef
                ts_sheet.cell(row=i, column=11).value = resultCdVer
            file.close
