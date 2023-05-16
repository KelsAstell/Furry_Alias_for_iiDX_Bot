import json
import os
import itertools as its
import pandas as pd

sourcepath = "Data.xlsx"
with open('fur_alias.json','r',encoding='utf-8') as fur:
    dict = json.load(fur)
    # print(dict[1])
excel_data = pd.read_excel(sourcepath)

# excel_data = pd.read_excel(sourcepath).to_dict()
# excel_data = pd.DataFrame(excel_data)
for i in range(len(excel_data)):
    flag = False
    # print(excel_data["Name"][i])
    for j in range(len(dict)):
        if dict[j]["Name"] == excel_data["Name"][i]:
            dict[j]["QQ"] = str(int(excel_data["QQ"][i]))
            flag = True
            if excel_data["Alias"][i] not in dict[j]["Alias"]:
                dict[j]["Alias"].append(excel_data["Alias"][i])
            if excel_data["Name"][i] not in dict[j]["Alias"]:
                dict[j]["Alias"].append(excel_data["Name"][i])
    if not flag:
        payload = {'ID': 1 + len(dict), 'QQ': str(int(excel_data["QQ"][i])),'Name': excel_data["Name"][i], 'Alias': [excel_data["Alias"][i]]}
        dict.append(payload)
json_str = json.dumps(dict, indent=4, ensure_ascii=False)
with open('fur_alias.json', 'w' ,encoding='utf-8') as fp:
    fp.write(json_str)
print('success')