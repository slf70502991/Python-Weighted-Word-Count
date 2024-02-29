import json
import os
import pandas as pd


print(os.getcwd())

file_to_analyse = input("Enter the json filename: ")
f = open(file_to_analyse, "r", encoding="utf-8")

data = json.load(f)
analyseLanguageParts = data['analyseLanguageParts'][0]
jobs = analyseLanguageParts['jobs'][0]

fileNames=[]
analyse = []
for job in analyseLanguageParts['jobs']:
    filename = job['fileName']
    fileNames.append(filename)
    analyse.append(job['data'])

contextMatch =[]
rep=[]
match100=[]
match95=[]
match85=[]
match75=[]
match50=[]
match0=[]
total = []
for i in analyse:
    contextMatch.append(i['contextMatch']['words'])
    rep.append(i['repetitions']['words'])
    match100.append(i['match100']['words']['sum'])
    match95.append(i['match95']['words']['sum'])
    match85.append(i['match85']['words']['sum'])
    match75.append(i['match75']['words']['sum'])
    match50.append(i['match50']['words']['sum'])
    match0.append(i['match0']['words']['sum'])
    total.append(i['total']['words'])

data_df= pd.DataFrame([fileNames, total ,contextMatch,rep, match100,match95,match85,match75,match50, match0], columns=fileNames)
data_df_transpose = data_df.T


# Calculate weighted word count 
def calculate_wc(data):
    wc_result=[]
    for index, rows in data_df_transpose.iterrows():
        # wc =  contextMatch*0.1 + rep * 0.1 +rep-match100*0.1 + match95 *0.2+ match85*0.3 + match75 + match50 + match0
        wc = rows[2] * 0.1+ rows[3] * 0.1 + rows[4] * 0.1 + rows[5] *0.2 + rows[6] *0.3 + rows[7] + rows[8] + rows[9]
        wc_result.append(wc)
    return wc_result

wc_result = calculate_wc(data_df_transpose)
data_df_transpose['Word Count'] = wc_result
data_df_transpose.to_excel('WWC.xlsx', engine="xlsxwriter")  
