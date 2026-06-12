import jsonata
import pandas as pd
import re

def strip(stripped):
    done = False
    while done == False:
        if len(stripped) > 0:
            if stripped[0]==" " or stripped[0]=="'" or stripped[0]=="[" or stripped[0]==",": stripped=stripped[1:]
            elif stripped[-1]==" " or stripped[-1]=="'" or stripped[-1]=="]": stripped=stripped[:-1]
            else:
                done = True
        else:
            done = True
    return stripped

def domainClItem(domain, var):
    return {"OID": "CL." + domain + ".DOMAIN",
            "Name": "Domain Abbreviation (" + domain + ")",
            "dataType": "text",
            "codeListType": "CodeListItem",
            "codeListItems": [{"code": domain,"codedValue": var}]} 


def get_ID(ID_string):
    if len(ID_string) <2: # if the ID string is None, return empty strings
        return "", ""
    else:
        o = 1 #letter it is looking at
        while ID_string[o] != ":" and o+1 < len(ID_string): #looking for the end of the ID
            o += 1
        if o == len(ID_string) - 1: # if the ID is not found, return empty strings
            ID_less = ID_string
            ID_less = strip(ID_less)
            return "", ID_less       
        else:
            Id = ID_string[1:o-1] # extracting the ID from the string
            ID_less = ID_string[o+1:]
            Id = strip(Id)
            ID_less = strip(ID_less)
            return Id, ID_less

def string_to_list(input, result):
    n = 0 #letter it is looking at
    while n < len(input) and input[n] != "}" and input[n] != "]": #looking for the end of the list
            if input[n-1:n+2] == ", '" or input[n-1:n+2] == ", """ or input[n] in ["{", "["]: # looking for the start of a new item in the list
               # print("n: ", n , "input[n-1:n+2]: ", input[n-1:n+2])
                n += 1
                m = n
                while m+1 < len(input) and input[m+1] not in ["}", "]"] and input[m:m+3] not in ("', '") and input[m:m+3] not in (""", """): # looking for the end of the item
                    m += 1
                Res=input[n:m+1] 
                result.append(Res) # appending the item to the list
                n = m + 1
            else: 
                n += 1
   # else:
       # result.append(input)

def string_to_list2(input, result): # new adapted string_to_list function (dec25)
    if input[0]=="{":
        n = 0 #letter it is looking at
        while input[n] != "}": #looking for the end of the list
            if input[n-1:n+2] == ", '" or input[n] in ["{"]: # looking for the start of a new item in the list
                n += 1
                m = n
                while m+1 < len(input) and input[m+1] not in ("}") and input[m:m+3] not in ("', '"): # looking for the end of the item
                    m += 1
                result.append(input[n:m+1]) # appending the item to the list
                n = m + 1
            else: 
                n += 1
    else:
        result.append(input)

def string_to_ID_list(input, result): # new adapted string_to_list function (dec25)
    if input[0]=="{":
        n = 0 #letter it is looking at
        while input[n] != "}": #looking for the end of the list
            if input[n-1:n+2] == ", '" or input[n] in ["{"]: # looking for the start of a new item in the list
                n += 1
                m = n
                while m+1 < len(input) and input[m+1] not in ("}") and input[m:m+3] not in ("', '"): # looking for the end of the item
                    m += 1
                Res=input[n:m+1] 
                id, Res = get_ID(Res)
                result[id]=Res # appending the item to the list
                n = m + 1
            else: 
                n += 1
    else:
        id, Res = get_ID(input)
        result[id]=Res

def string_to_nested_list(input, resultarm, result):
    n = 0 #letter it is looking at
    if input[0] != "[":
        input = "[" + input + "]"
    while input[n] != "]": #looking for the end of the list
        if input[n] == "{": # looking for the start of a new item in the list
            n += 1
            m = n
            while m+1 < len(input) and input[m:m+3] not in ("', '"):
                m += 1
            resultarm.append(input[n:m+1]) # appending the item to the list
            if resultarm[-1][-1]==",": resultarm[-1]=resultarm[-1][:-1] # remove trailing comma if it exists
            n = m + 1
            while m+1 < len(input) and input[m+1] != "}": # looking for the end of the item
                m += 1
            result.append(input[n:m+1]) # appending the item to the list
            n = m + 1
        else:
            n += 1
                
def Parse_jsonata(codeSnip,data):
        if codeSnip is None:
            result = " "
        else:
            try:
                expr = jsonata.Jsonata(codeSnip)
                result = expr.evaluate(data)  
            except:
                result = "Error in expression " + codeSnip
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
            if result0.count("', '") == 0:
                result0 = result0[1:-1]
                result0 = result0.replace("}, {", ", ")
        return result0

