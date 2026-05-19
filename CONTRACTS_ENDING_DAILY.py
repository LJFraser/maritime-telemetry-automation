# -*- coding: utf-8 -*-
"""
Created on Mon Nov 27 10:19:30 2023
@author: Leane Fraser
"""
import psycopg2
import psycopg2.extras
from datetime import datetime, timedelta
import win32com.client as win32             # pip install pypiwin32 (if not already installed)
import pandas as pd
import sys

print("-------------------------------------------------------")
print("|                                                     |")
print("|                CONTRACTS ENDING TODAY               |")
print("|                                                     |")
print("-------------------------------------------------------\n")

flyn_positions_host = "your-rds-endpoint.eu-west-3.rds.amazonaws.com"
read_db_username = "your_read_only_user"
db_password = "your_secure_password"
positions_db_name = "YourDatabaseName"

# Establish connection using generic variables
conn = psycopg2.connect(
    host=flyn_positions_host,
    database=positions_db_name,
    user=read_db_username,
    password=db_password,
)

logo_absolute = str("C:\project_data\templates\logo.jpg")

now = datetime.now()
today = now.strftime("%Y-%m-%d")

tomorrow = now + timedelta(days = 1)
tomorrow = tomorrow.strftime("%Y-%m-%d")
year = datetime.now().strftime("%Y")
sql = f"""SELECT producttype,type,CAST(customerid AS varchar),id,name,mmsi,yachttype,yachtname,predeliverydateutc,startdateutc,enddateutc,postdeliverydateutc
FROM "FLYNContractsData"."FLYNContracts" WHERE enddateutc > '{today}' AND enddateutc < '{tomorrow}' AND status IN ('ACTIVE','DELIVERED','CLOSED')
OR postdeliverydateutc > '{today}' AND postdeliverydateutc < '{tomorrow}' AND status IN ('ACTIVE','DELIVERED','CLOSED')
ORDER BY producttype, customerid ;"""

cursor.execute(sql)
b = cursor.fetchall()
df = pd.DataFrame.from_dict(b)
today_hrs =f'{today} 23:59:59'

try:
    df2 = df.drop(df[df['postdeliverydateutc'] > today_hrs].index)
    df2.reset_index(drop=True, inplace=True)
    nb_ctrs_end_td = len(df2)
except:
    print(f"No PREMIUM or REGULAR contracts ending today ({today})")
    sys.exit()

no_pr = len(df2[df2.producttype.str.contains('PREMIUM')])
premium = df2[df2.producttype.str.contains('PREMIUM')]
premium.reset_index(drop=True, inplace=True)

if len(premium) > 0:

    premium.to_excel(f'C:\project_data\templates\{today}_contracts_ending_{no_pr}.xlsx', index = False, header = True)
    print(f"{no_pr} PREMIUM contract(s) ending today.\n")
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = 'operations-team1@example.com'
#    mail.To = 'operations-team1@example.com;operations-team2@example.com;operations-team3@example.com'     #activate when building .exe
    mail.Subject = f'{no_pr} PREMIUM CONTRACT(S) ENDING TODAY {today} - automatic email'
    mail.HTMLBody += """<h1 style="font-family: 'Calibri'; font-size: 23; font-weight: bold; color: #0A506E;">PREMIUM CONTRACTS ENDING TODAY - please issue the report and create the certificate.</h1>"""

    for contract in range (0,no_pr):
        mail.HTMLBody += f"""<h4 style="font-family: 'Calibri'; font-size: 14; color: #080909;">Customer No: {premium['customerid'][contract]} - {premium['type'][contract]} - {premium['yachttype'][contract]}_{premium['yachtname'][contract]}_{premium['mmsi'][contract]}_{premium['id'][contract]} | Contract end: {premium['enddateutc'][contract]} |  Post delivery: {premium['postdeliverydateutc'][contract]}</h4>"""


    attachment  = f"C:\project_data\templates\{today}_contracts_ending_{no_pr}.xlsx"
    mail.Attachments.Add(attachment)
    mail.Attachments.Add(logo_absolute)
    image = mail.Attachments.Add(logo_absolute)
    mail.HTMLBody += """ <div><br><img src="cid:logo-img" width=5%></div>"""
    image.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "logo-img")
    mail.Send()

regular = df2[df2.producttype.str.contains('REGULAR')]
regular.reset_index(drop=True, inplace=True)
regular.replace('1', 'SOS Yachting France', inplace=True)
regular.replace('2', 'SOS Yachting Italy', inplace=True)
no_re = len(regular)
no_sosFR = len(regular[regular.customerid.str.contains('SOS Yachting France')])
no_sosIT = len(regular[regular.customerid.str.contains('SOS Yachting Italy')])

if len(regular) > 0:

    regular.to_excel(f'C:\project_data\templates\{today}_contracts_ending_{no_re}.xlsx', index = False, header = True)
    print(f"{no_re} REGULAR contracts ending today: \n\n{no_sosFR} for SOS France  \n{no_sosIT} for SOS Italy  \n\nnow sending automated email(s).")
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = 'operations-team1@example.com'
#   mail.To = 'operations-team1@example.com;operations-team2@example.com;operations-team3@example.com'     #activate when building .exe
    mail.Subject = f'{no_re} REGULAR CONTRACT(S) ENDING {today} | SOS France: {no_sosFR} | SOS Italy: {no_sosIT} - automatic email'
    mail.HTMLBody += """<h1 style="font-family: 'Calibri'; font-size: 23; font-weight: bold; color: #0A506E;">REGULAR CONTRACTS ENDING TODAY - please send auto reports.</h1>"""

    for contract in range (0,no_re):
        mail.HTMLBody += f"""<h4 style="font-family: 'Calibri'; font-size: 14; color: #080909;">{regular['customerid'][contract]} |
        {regular['type'][contract]} - {regular['yachttype'][contract]}_{regular['yachtname'][contract]}_{regular['mmsi'][contract]}_{regular['id'][contract]}
        | Contract end: {regular['enddateutc'][contract]} |  Post delivery: {regular['postdeliverydateutc'][contract]}</h4>"""

    attachment  = f"C:\project_data\templates\{today}_contracts_ending_{no_re}.xlsx"
    mail.Attachments.Add(attachment)
    mail.Attachments.Add(logo_absolute)
    image = mail.Attachments.Add(logo_absolute)
    mail.HTMLBody += """ <div><br><img src="cid:logo-img" width=5%></div>"""
    image.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "logo-img")
    mail.Send()
