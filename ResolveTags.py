import jsonata
import re
import pandas as pd


def ResolveTag(Txt,data):
# <usdm:tag name="min_age"/>
   #print("Resolving tag in text: ", Txt)
   while Txt.find("<usdm:tag") != -1:
        m = re.search(r'.*<usdm:tag name="([^"]*)"/>', Txt, re.DOTALL | re.IGNORECASE)
        if m:
            #print("Found tag: ", m.group(1))
            attrs = m.group(1)
            # print (attrs, m.end(0), m.start(1))
            NewTxt=Get_TagValue(attrs,data)
            Txt2=Txt[0:m.start(1)-16] + str(NewTxt) + Txt[m.end(0):len(Txt)]
            # print(Txt2)
            Txt=Txt2
   result=Get_plainText(Txt)
   return result   


def Get_plainText(TxtRich):
    # replace <li> with -
    TxtRich = re.sub(r'<li>(.*?)</li>', r' - \1;', TxtRich, flags=re.DOTALL | re.IGNORECASE)
    # replace <br> with space
    TxtRich = re.sub(r'<br\s*/?>', ' ', TxtRich, flags=re.DOTALL | re.IGNORECASE)
    # replace <tr> with space
    TxtRich = re.sub(r'<tr\s*/?>', ' ', TxtRich, flags=re.DOTALL | re.IGNORECASE)
    # remove <p> and </p>
    TxtRich = re.sub(r'</?p\s*/?>', '', TxtRich, flags=re.DOTALL | re.IGNORECASE)
    #remove <ol> and </ol>
    TxtRich = re.sub(r'</?ol\s*/?>', '', TxtRich, flags=re.DOTALL | re.IGNORECASE)
    #remove <ul> and </ul>
    TxtRich = re.sub(r'</?ul\s*/?>', '', TxtRich, flags=re.DOTALL | re.IGNORECASE)
    # remove all other HTML tags
    TxtRich = re.sub(r'<[^>]+>', '', TxtRich, flags=re.DOTALL | re.IGNORECASE)
    # replace &lt; with <
    TxtRich = TxtRich.replace('&lt;', '<')  
    TxtRich = TxtRich.replace('&le;', '<=')  
    # replace &gt; with >
    TxtRich = TxtRich.replace('&gt;', '>')  
    # replace \' with '
    TxtRich = TxtRich.replace('\'', "'")
    # replace &gt; with >=
    TxtRich = TxtRich.replace('&gt;', '>=')  
    # replace &amp; with &
    TxtRich = TxtRich.replace('&amp;', '&')     
    # replace &#174; with ®     
    TxtRich = TxtRich.replace('&#174;', ' (R)')     
    # replace &#8482; with ™    
    TxtRich = TxtRich.replace('&#8482;', ' (TM)')     
    # replace &#169; with © 
    TxtRich = TxtRich.replace('&#169;', ' (C)')
    # replace ≤ with <=
    TxtRich = TxtRich.replace('≤', '<=')

    TxtRich = TxtRich.replace('&#181;', 'micro')
    
    # collapse whitespace
    TxtRich = re.sub(r'\s+', ' ', TxtRich).strip()
   
    
    TxtRich = TxtRich.replace('\\n-', '-')
    TxtRich = TxtRich.replace('\n-', '-')
    # remove \n
    TxtRich = TxtRich.replace("\\n", "")
    # remove multiple spaces    
    TxtRich = re.sub(' +', ' ', TxtRich)    

    TxtRich = re.sub('\\n', '', TxtRich)        
    TxtRich = re.sub('\n', '', TxtRich)   
    
    TxtRich = TxtRich.replace(";", "; ") 

    
    # collapse whitespace
    TxtRich = re.sub(r'\s+', ' ', TxtRich).strip()
   
    # print(TxtRich)
    return TxtRich

def Get_TagValue(tag,data):
    jsonataString = "study.versions.dictionaries.parameterMaps[tag='" + tag + "'].reference"
    expr = jsonata.Jsonata(jsonataString)
    reference = expr.evaluate(data)
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
                jsonataString3 = "study.**[instanceType='" + klass + "'  and id='" + id + "']." + attr
                # print(jsonataString2)
                expr2 = jsonata.Jsonata(jsonataString3)
                value = expr2.evaluate(data)
                # print("value: ", value)
            except:
                value="//TAG REFERENCE PARSING ERROR//"
    return value