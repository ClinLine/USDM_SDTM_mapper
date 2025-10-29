import jsonata
import json
import openpyxl
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

def string_to_nested_list(input, resultarm, result):
    n = 0 #letter it is looking at
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

def ResolveTag(Txt,data):
# <usdm:tag name="min_age"/>
   result=Txt
   m = re.search(r'.*<usdm:tag name="([^"]*)"/>', Txt, re.DOTALL | re.IGNORECASE)
   if m:
        attrs = m.group(1)
        # print (attrs, m.end(0), m.start(1))
        NewTxt=Get_TagLocation(attrs,data)
        Txt2=Txt[0:m.start(1)-16] + str(NewTxt) + Txt[m.end(0):len(Txt)]
        # print(Txt2)
        result=Txt2
   return result   


def Get_TagLocation(tag,data):
    jsonataString = "study.versions.dictionaries.parameterMaps[tag='" + tag + "'].reference"
    expr = jsonata.Jsonata(jsonataString)
    reference = expr.evaluate(data)
    # print("location : ", reference)
    if reference is None:
        value="//TAG NOT IN DICTIONARY//"
    else:
        if reference[0] != "<": 
            value=reference
        else:
            try:
                m = re.search(r'.*klass="([^"]*)"', reference, re.DOTALL | re.IGNORECASE)
                klass = m.group(1)
                m = re.search(r'.*id="([^"]*)"', reference, re.DOTALL | re.IGNORECASE)
                id = m.group(1)
                m = re.search(r'.*attribute="([^"]*)"', reference, re.DOTALL | re.IGNORECASE)
                attr = m.group(1)
                jsonataString2 = "study.versions" + ClassToRelation(klass) + "[id='" + id + "']." + attr
                # print(jsonataString2)
                expr2 = jsonata.Jsonata(jsonataString2)
                value = expr2.evaluate(data)
                # print("value: ", value)
            except:
                value="//TAG REFERENCE PARSING ERROR//"
    return value
   
def ClassToRelation(klass):
    if klass == "Activity": return ".studyDesigns.activities"
    elif klass == "Quantity": return "..."   
    elif klass == "Indication": return ".studyDesigns.indications"
    elif klass == "Objective": return ".studyDesigns.objectives"
    elif klass == "Endpoints": return ".studyDesigns.objectives.endpoints"
    elif klass == "StudyDesignPopulation": return ".studyDesigns.population"
    elif klass == "BiomedicalConcept": return ".biomedicalConcepts"
    elif klass == "BiomedicalConceptProperty": return ".biomedicalConcepts.properties"
    elif klass == "StudyIntervention": return ".studyInterventions"
    elif klass == "AdministrableProduct": return ".administrableProducts"
    elif klass == "ResponseCode": return "biomedicalConcepts.properties.responseCodes"
    elif klass == "MedicalDevice": return ".medicalDevices"
    elif klass == "StudyRole": return ".roles"
    else: return None
