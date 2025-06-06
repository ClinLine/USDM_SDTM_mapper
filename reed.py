import jsonata
import json
import pandas as pd
with open('ReCoPad.json', 'r') as file:
    data = json.load(file)

expr = jsonata.Jsonata("study.versions.studyDesigns.blindingSchema.standardCode.decode")
result = expr.evaluate(data)
print(result)

df = pd.read_excel('your_file.xlsx')
print(df)