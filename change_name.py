import os
import pandas as pd
import json

path_excel = os.getcwd() + "/uvjeti_certifikat_2022.xlsx"

def change_name(email, team_name, full_name):
    if os.path.isdir('./' + 'Certifikati-2022/' + team_name + '/' + email):
      if os.path.isfile('./' + 'Certifikati-2022/' + team_name + '/' + email+ '/' + full_name + ' eSTUDENT Certifikat 2020-2022.pdf'):
        os.rename('./Certifikati-2022'+ '/' + team_name + '/' + email + '/' + full_name + ' eSTUDENT Certifikat 2020-2022.pdf', './Certifikati-2022'+ '/' + team_name + '/' + email + '/' + full_name + ' eSTUDENT Certifikat 2021-2022.pdf')
    else:
        return False

def get_data_from_excel(key):
    excel_data_dict = {}
    table = pd.read_excel(path_excel)
    excel_data_dict["team_name"] = table.loc[key][4]
    print(table.loc[key][1])
    excel_data_dict["position"] = table.loc[key][5]
    excel_data_dict["email"] = table.loc[key][2]
    excel_data_dict["full_name"] = table.loc[key][0] + table.loc[key][1]
    #excel_data_dict["email"] = "iva.krizanac@estudent.hr"
    return excel_data_dict

excel_table = pd.read_excel(path_excel).to_json()
excel_table = json.loads(excel_table)
for key in excel_table["Ime"]:
    key = int(key)
    new_key = key
    get_data_from_excel(new_key)
    excel_data_dict = get_data_from_excel(key)
    email = excel_data_dict["email"]
    team_name = excel_data_dict["team_name"]
    full_name = excel_data_dict["full_name"]
    change_name(email,team_name,full_name)

