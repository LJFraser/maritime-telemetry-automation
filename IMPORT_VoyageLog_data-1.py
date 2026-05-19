# -*- coding: utf-8 -*-
"""
Created on Thu Aug 10 10:34:44 2023

@author: Leane Fraser
"""
import pandas as pd
from colorama import Back
from colorama import Style
import re
import os

print("-------------------------------------------------------")
print("|                                                     |")
print("|            REFORMAT & SPLIT A VoyageLog FILE        |")
print("|                                                     |")
print("-------------------------------------------------------\n")

filename = input("Enter the name of file (must be in C:\project_data\FLYN_file_input) to be reformatted for database ingestion: \n")

posfile = pd.read_csv(f'C:\project_data\FLYN_file_input\{filename}', delimiter=',')
pos_file_nonans = posfile.dropna(how='all')
pos_file_nonans.columns = ['No','Date','Time','Source','Type','LAT','LON','S.SRC','SOG/kn','COG/deg T','S.SRC','HDG/deg T','S.SRC','Corr/deg T','Wind/kn','Wind/deg T','Dist/NM','Depth/m','Description']
pos_file_nonans['TIMESTAMP_UTC'] = pos_file_nonans['Date'] + ' ' + pos_file_nonans['Time']
digits = re.split('_', filename)
mmsi = digits[2]
dsrc = 'CAPT'

pos_file_nonans['MMSI'] = mmsi
pos_file_nonans['DSRC'] = dsrc

pos_file_nonans['LAT_DEGREE'] = pos_file_nonans.LAT.astype(str).str[:2]
pos_file_nonans['LAT_DECIMAL_M'] = pos_file_nonans.LAT.astype(str).str[3:9]
pos_file_nonans['LON_DEGREE'] = pos_file_nonans.LON.astype(str).str[:3]
pos_file_nonans['LON_DECIMAL_M'] = pos_file_nonans.LON.astype(str).str[4:10]

a = (pos_file_nonans['LAT_DEGREE'])
b = pd.to_numeric(a, downcast='signed')
c = (pos_file_nonans['LAT_DECIMAL_M'])
d = pd.to_numeric(c, downcast='signed')
#FINAL COLUMN TO EXPORT IN GOOD FORMAT
lat_decimal_deg = b + (d/60)

a1 = (pos_file_nonans['LON_DEGREE'])
b1 = pd.to_numeric(a1, downcast='signed')
c1 = (pos_file_nonans['LON_DECIMAL_M'])
d1 = pd.to_numeric(c1, downcast='signed')
#FINAL COLUMN TO EXPORT IN GOOD FORMAT
lon_decimal_deg = b1 + (d1/60)

pos_file_nonans['LATITUDE'] = lat_decimal_deg
pos_file_nonans['LONGITUDE'] = lon_decimal_deg

name=str(digits[0])
to_export = pos_file_nonans[['MMSI', 'TIMESTAMP_UTC','LATITUDE','LONGITUDE','DSRC']]
rows = len(to_export)

print(f"\nThere are {Back.CYAN}{rows}{Style.RESET_ALL} positions in this file.")
del_interval = input("\nEnter the number of positions to be added per file: ")
print("\nThe following files have been extracted to C:\project_data\FLYN_file_input: \n")

int1 = 0
int2 = int(del_interval)
del_interval = int(del_interval)

while int2 < rows + del_interval:
    extract = (to_export.iloc[int1:int2,0:])
    extract.to_csv(f'C:\project_data\FLYN_file_input\\{digits[0]}_{mmsi}_POSITIONS_{int1}_{int2}_captain_extraction_date_{digits[8]}-{digits[9]}-{digits[10][0:2]}.csv', index = False)
    print(f"{digits[0]}_{mmsi}_POSITIONS_{int1}_{int2}_captain_extraction_date_{digits[8]}-{digits[9]}-{digits[10][0:2]}.csv")
    int1 = int1 + del_interval
    int2 = int2 + del_interval

remove = input(f"\n{Back.CYAN}CLEAN UP OPTION:{Style.RESET_ALL} \nWould you like to delete the original input file? - {Back.GREEN} Y {Style.RESET_ALL} or {Back.RED} N {Style.RESET_ALL}\n")

if remove == 'Y' or remove == 'y' or remove == 'yes' or remove == 'YES' or remove == 'Yes':
    os.remove(f'G:\\_FLYN\\FLYN_file_input\\{filename}')
    print(f"\nFile {filename} has been deleted.")
elif remove == 'N' or remove == 'n' or remove == 'no' or remove == 'NO' or remove == 'No':
    print(f"\nOK file {filename} has been kept.")

print("\nYou are now ready to run FLYN_6 on the above output files.")
