# -*- coding: utf-8 -*-
"""
Created on Tue Sep 19 11:29:17 2023

@author: Leane Fraser
"""
import psycopg2
import psycopg2.extras
from datetime import datetime, timedelta
import os
import shutil
import pandas as pd
import sys
#import emoji                               # pip install emoji (if not already installed)
import win32com.client as win32             # pip install pypiwin32 (if not already installed)
from colorama import init as colorama_init  # pip install colorama
from colorama import Back
from colorama import Style
from subprocess import call

import logging

logging.basicConfig(filename='G:\_FLYN\FLYN_log\FLYN_CERTIFICATE.log', encoding='utf-8', level=logging.DEBUG)
logger = logging.getLogger()
logger.setLevel(logging.DEBUG)
now_dtr = datetime.utcnow()
now_dtg = datetime.strftime(now_dtr, "%Y-%m-%d %H:%M:%S")
logger.info('\n' + now_dtg + 'Z===========================')
logger.info('| CREATION OF A CERTIFICATE |')

#-------------------------CONNECTION TO DATABASE-------------------------------
import os

conn = psycopg2.connect(
    host=os.getenv("DB_HOST"),
    database=os.getenv("DB_NAME"),
    user=os.getenv("DB_USER"),
    password=os.getenv("DB_PASSWORD"),
)
cursor = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)

print("-------------------------------------------------------")
print("|                                                     |")
print("|            SCRIPT TO MAKE A CERTIFICATE             |")
print("|                                                     |")
print("-------------------------------------------------------\n")
print("You can choose from the following options:\n")
print("1. CHARTER certificate. \n2. TRANSPORT certificate. \n3. INTERNATIONAL SALE certificate.\n")
option = input('Please enter a number between 1-3:\n')

#------------------------------- DEFINED FUNCTIONS -----------------------------

year = datetime.now().strftime("%Y")

root = 'C:\project_data\templates'

def timezone_change(timezone, cert, ctype):
        logger.info(f'Timezone is {timezone}')

        if timezone == 'CEST':
            logger.info('2 hours have been added to obtain local time.')

            if ctype == 'INTERNATIONAL_SALE':
                cert[0]['Enter_dtg_local'] = pos_intwat[0]['datetime'] + timedelta(hours=2)
                cert[0]['Exit_dtg_local'] = pos_intwat[1]['datetime'] + timedelta(hours=2)
                cert[0]['approx_end_local_txt'] = cert[0]['approx_end_utc'] + timedelta(hours=2)
                cert[0]['approx_start_local_txt'] = cert[0]['approx_start_utc'] + timedelta(hours=2)
            else:
                cert[0]['start_local'] = cert[0]['startdateutc'] + timedelta(hours=2)
                cert[0]['end_local'] = cert[0]['enddateutc'] + timedelta(hours=2)

        if timezone == 'CET':

            logger.info('1 hour has been added to obtain local time.')

            if ctype == 'INTERNATIONAL_SALE':
                cert[0]['Enter_dtg_local'] = pos_intwat[0]['datetime'] + timedelta(hours=1)
                cert[0]['Exit_dtg_local'] = pos_intwat[1]['datetime'] + timedelta(hours=1)
                cert[0]['approx_end_local_txt'] = cert[0]['approx_end_utc'] + timedelta(hours=1)
                cert[0]['approx_start_local_txt'] = cert[0]['approx_start_utc'] + timedelta(hours=1)

            else:
                cert[0]['start_local'] = cert[0]['startdateutc'] + timedelta(hours=1)
                cert[0]['end_local'] = cert[0]['enddateutc'] + timedelta(hours=1)

        if timezone == 'UTC':

            logger.info('No change to the times - all UTC.')

            if ctype == 'INTERNATIONAL_SALE':
                cert[0]['Enter_dtg_local'] = pos_intwat[0]['datetime']
                cert[0]['Exit_dtg_local'] = pos_intwat[1]['datetime']
                cert[0]['approx_end_local_txt'] = cert[0]['approx_end_utc']
                cert[0]['approx_start_local_txt'] = cert[0]['approx_start_utc']
            else:
                cert[0]['start_local'] = cert[0]['startdateutc']
                cert[0]['end_local'] = cert[0]['enddateutc']

        if timezone == 'AST':
            logger.info('4 hours have been subtracted to obtain local time.')

            if ctype == 'INTERNATIONAL_SALE':
                cert[0]['Enter_dtg_local'] = pos_intwat[0]['datetime'] - timedelta(hours=4)
                cert[0]['Exit_dtg_local'] = pos_intwat[1]['datetime'] - timedelta(hours=4)
                cert[0]['approx_end_local_txt'] = cert[0]['approx_end_utc'] - timedelta(hours=4)
                cert[0]['approx_start_local_txt'] = cert[0]['approx_start_utc'] - timedelta(hours=4)
            else:
                cert[0]['start_local'] = cert[0]['startdateutc'] - timedelta(hours=4)
                cert[0]['end_local'] = cert[0]['enddateutc'] - timedelta(hours=4)

def fiscal_status(ctype, cert):

    print("\nYou are initialising the certificate for: " + str(cert[0]['cname']))
    print('\nThis is a ' + str(ctype) +' contract.')

    if cert[0]['status'] != "POST_DELIVERY ENDS" and cert[0]['status'] != "CRUISE ENDS":
        print("")
#       print(emoji.emojize(":thumbs_down:\n"))
        print("WARNING - the status needs to be on \"CRUISE ENDS or POST_DELIVERY ENDS\" before issuing the Auto Report and Certificate.\n")
        logger.info("WARNING - the status is not \"CRUISE ENDS or POST_DELIVERY ENDS\" for this contract.\n")
        print("The current status for " + cert[0]['cname'] + " is: ")
        print(cert[0]['cname'])
        print("\nSCRIPT HAS ENDED - no folder created, please try again later.")
        sys.exit()

    else:
#       print(emoji.emojize("\n:grinning_face:\n"))
        print("\nThe status for " + cert[0]['cname'] + " is on CRUISE ENDS or POST_DELIVERY ENDS; process to make the certificate inputs has been initiated.\n")
        logger.info("Status for " + cert[0]['cname'] + " is on CRUISE ENDS or POST_DELIVERY ENDS; process to make the certificate initiated.")

def trip_calculations(cert):

    cert[0]['nb_tot'] = cert[0]['nb_pre'] + cert[0]['nb_post'] + cert[0]['nb_ccial']
    cert[0]['nb_pre'] = ('no') if cert[0]['nb_pre'] == 0 else cert[0]['nb_pre']
    cert[0]['nb_pre_hiseas'] = ('no') if cert[0]['nb_pre_hiseas'] == 0 else cert[0]['nb_pre_hiseas']
    cert[0]['nb_post'] = ('no') if cert[0]['nb_post'] == 0 else cert[0]['nb_post']
    cert[0]['nb_post_hiseas'] = ('no') if cert[0]['nb_post_hiseas'] == 0 else cert[0]['nb_post_hiseas']
    cert[0]['nb_emb'] = ('no') if cert[0]['reference'] == 'STATIC' or cert[0]['nb_emb'] == 0 else cert[0]['nb_emb']
    cert[0]['nb_ccial'] = ('no') if cert[0]['nb_ccial'] == 0 or cert[0]['reference'] == 'STATIC' else cert[0]['nb_ccial']

def change_name(cert):

    cert[0]['yachtname'] = cert[0]['yachtname'][0:4] if 'LOON' in cert[0]['yachtname'] else cert[0]['yachtname']

def year_path(year, ctype):

    if ctype == 'CHARTER':

        path = root + str(year) + '\__certificate_making_CH'
        certificate = root + str(int(year)-1) + '\__certificate_making_CH\FLYN_Certificate_CHARTER xxx under # iii_aaaa-mm-dd.docx'
        destination = path + '\FLYN_Certificate_CHARTER xxx under # iii_aaaa-mm-dd.docx'

        isExist = os.path.exists(path)
        if isExist == False:
                os.makedirs(path)
                print('\nNEW YEAR PATH: ' +path + '  has been created')
                logger.info("Path created for " + path)
                shutil.copy(certificate, destination)
                print("\nTemplate for certificate copied into directory.")


    elif ctype == 'TRANSPORT':

        path2 = root + str(year) + '\__certificate_making_TR'
        certificate = root + str(int(year)-1) + '\__certificate_making_TR\FLYN_Certificate_TRANSPORT xxx under # iii_aaaa-mm-dd.docx'
        destination = path2 + '\FLYN_Certificate_TRANSPORT xxx under # iii_aaaa-mm-dd.docx'

        isExist = os.path.exists(path2)
        if isExist == False:
                os.makedirs(path2)
                print('\nNEW YEAR PATH: ' + path2 + '  has been created')
                logger.info("Path created for " + path2)
                shutil.copy(certificate, destination)
                print("Template for certificate copied into directory.")

    else:
        path3 = root + str(year) + '\__certificate_making_international_sale'
        certificate = root + str(int(year)-1) + '\__certificate_making_international_sale\FLYN_Certificate_INTERNATIONAL_SALE xxx under # iii_aaaa-mm-dd.docx'
        destination = path3 + '\FLYN_Certificate_INTERNATIONAL_SALE xxx under # iii_aaaa-mm-dd.docx'

        isExist = os.path.exists(path3)
        if isExist == False:
                os.makedirs(path3)
                print('\nNEW YEAR PATH: ' + path3 + '  has been created')
                logger.info("Path created for " + path3)
                shutil.copy(certificate, destination)
                print("Template for certificate copied into directory.")

def folder_creation(ctype, cert):

    now=datetime.now()
    modified = now + timedelta(minutes = 30)
    cert[0]['modified'] = modified.strftime("%Y-%m-%d %H:%M")
    cert[0]['date_cert'] = now.strftime("%Y-%m-%d")

    isExist = os.path.exists(path)

    if isExist == False:
        os.mkdir(path)
        print("\nThe folder: " + str(foldername) + " has been created under " + path[:length]  + "\n")
        logger.info("The folder: " + str(foldername) + " has been created under " + path[:length])
        print("containing the following files:\n")
        print(" "*30,"FINAL_mailmerge_INPUTS.xlsx")


        if ctype == 'INTERNATIONAL_SALE':
            print(" "*30,'Certificate_for_the_INTERNATIONAL_SALE_in_'+ c[0]['year'] + '_of_'  + c[0]['yachtname'] + '.docx')
            print(" "*30, 'Positions_Table.xlsx')

        else:
            print(" "*30,"FLYN_Certificate_" + str(ctype) +"_" +str(foldername) + '_' + cert[0]['date_cert'] + '.docx')

    else:
        print("\nThe folder: " + str(foldername) + " already exists.\n")
        logger.info("The folder: " + str(foldername) + " already exists.")
        print("This folder contains the following files:\n")
        print(" "*30, "FINAL_mailmerge_INPUTS.xlsx")

        if ctype == 'INTERNATIONAL_SALE':
            print(" "*30,'Certificate_for_the_INTERNATIONAL_SALE_in_'+ cert[0]['year'] + '_of_' + cert[0]['yachtname'] + '.docx')
            print(" "*30, 'Positions_Table.xlsx')

        else:
            print(" "*30,"FLYN_Certificate_" + str(ctype) +"_" +str(foldername) + '_' + cert[0]['date_cert'] + '.docx')

    if ctype == 'INTERNATIONAL_SALE':

        df_final.to_excel(path + 'Positions_Table.xlsx', index=False, header=True)
        certificate = path[:length] + 'FLYN_Certificate_INTERNATIONAL_SALE xxx under # iii_aaaa-mm-dd.docx'
        destination = path + 'Certificate_for_the_INTERNATIONAL_SALE_in_' + cert[0]['year'] + '_of_' + cert[0]['yachtname'] + '.docx'

    else:
        certificate = path[:length] + 'FLYN_''Certificate_' + ctype + ' xxx under # iii_aaaa-mm-dd.docx'
        destination = path + '\FLYN_Certificate_' + ctype + '_' + str(foldername) + '_' + cert[0]['date_cert'] + '.docx'

    shutil.copy(certificate, destination)
    df = pd.DataFrame(cert)
    df.to_excel(path + '\FINAL_mailmerge_inputs.xlsx', index=False, header=True)

def mail_merge(ctype, cert):

    try:

        print("\nMICROSOFT OFFICE WORD is loading ... please click OK and preview results of the Mail Merge.")
        working_directory = path
        source_name = 'FINAL_mailmerge_INPUTS.xlsx'
        WordApp = win32.Dispatch('Word.Application')
        WordApp.Visible = True

        if ctype == 'INTERNATIONAL_SALE':
            sourceDoc = WordApp.Documents.Open(os.path.join(working_directory, 'Certificate_for_the_INTERNATIONAL_SALE_in_'+ c[0]['year'] + '_of_' + c[0]['yachtname'] + '.docx'))
        else:
            sourceDoc = WordApp.Documents.Open(os.path.join(working_directory, 'FLYN_Certificate_' + ctype + '_' +str(foldername) + '_' +  cert[0]['date_cert'] + '.docx'))

        mail_merge = sourceDoc.MailMerge
        mail_merge.OpenDataSource(Name:=os.path.join(working_directory, source_name),sqlstatement:="SELECT * FROM [Data Source$]")
        logger.info("Microsoft Office Word opened with Mail Merge application initiated.")

    except:
        print('error: ', str(sys.exc_info()))
        logger.info('Check if Microsoft Office Word is already open.')

def manual_positions():

        pos_intwat[0]['datetime'] = input(f"\n\nPlease enter manually the {Back.GREEN}DATE & TIME (Z) of ENTRY{Style.RESET_ALL} into international waters using the format: {Back.GREEN}YYYY-MM-DD HH:mm\n{Style.RESET_ALL}")
        pos_intwat[0]['lon'] = input(f"Please enter manually the {Back.CYAN}ENTRY LONGITUDE{Style.RESET_ALL} coordinates in decimal degrees e.g. 7.465641{Back.CYAN}\n{Style.RESET_ALL}")
        pos_intwat[0]['lat'] = input(f"Please enter manually the {Back.CYAN}ENTRY LATITUDE{Style.RESET_ALL} coordinates in decimal degrees e.g. 43.495359{Back.CYAN}\n{Style.RESET_ALL}")
        pos_intwat[1]['datetime'] = input(f"\nPlease enter manually the {Back.RED}DATE & TIME (Z) of EXIT{Style.RESET_ALL} out of international waters using the format: {Back.RED}YYYY-MM-DD HH:mm\n{Style.RESET_ALL}")
        pos_intwat[1]['lon'] = input(f"Please enter manually {Back.CYAN}EXIT LONGITUDE{Style.RESET_ALL} coordinates in decimal degrees e.g. 7.465641{Back.CYAN}\n{Style.RESET_ALL}")
        pos_intwat[1]['lat'] = input(f"Please enter manually the {Back.CYAN}EXIT LATITUDE{Style.RESET_ALL} coordinates in decimal degrees e.g. 43.495359{Back.CYAN}\n{Style.RESET_ALL}")

        print("\n\nYou have entered the following:")
        print("Entry time:", pos_intwat[0]['datetime'], " | Longitude:", pos_intwat[0]['lon'], "| Latitude:", pos_intwat[0]['lat'])
        print("Exit time: ", pos_intwat[1]['datetime'], " | Longitude:", pos_intwat[1]['lon'], "| Latitude:", pos_intwat[1]['lat'])
        print("\nPositions have been converted into DECIMAL DEGREES for the certificate:\n")

        pos_intwat[0]['datetime'] = datetime.strptime(pos_intwat[0]['datetime'], '%Y-%m-%d %H:%M')
        pos_intwat[1]['datetime'] = datetime.strptime(pos_intwat[1]['datetime'], '%Y-%m-%d %H:%M')
        pos_intwat[0]['lon'] = float(pos_intwat[0]['lon'])
        pos_intwat[0]['lat'] = float(pos_intwat[0]['lat'])
        pos_intwat[1]['lon'] = float(pos_intwat[1]['lon'])
        pos_intwat[1]['lat'] = float(pos_intwat[1]['lat'])

        logger.info("Auto positions not accepted. Positions have been manually entered by user.")
        c[0]['EXIT_TW'] = str('YES')

def call_qgis(cert):
    print(f'{Back.RED}\nPLEASE CHECK THE MAP OF{Style.RESET_ALL} ' + cert[0]['cname'] + f' {Back.RED}WITH QGIS TO CONFIRM:{Style.RESET_ALL}\n')
    print("QGIS is currently loading ... please check the map to confirm positions and save images, then CLOSE QGIS to continue with script.")
    call(r"C:\Program Files\QGIS 3.22.16\bin\qgis-ltr-bin.exe")
    logger.info('QGIS was opened by script.')

#------------------------------------------------------------------------------
#                               CHARTER CERTIFICATE
#------------------------------------------------------------------------------
if option == '1':

    year_path(year = year, ctype = 'CHARTER')
    logger.info('CHARTER certificate initiated.')
    contractid = input('Enter yacht contract number: ')
    logger.info('Contract input: ' + str(contractid))

    try:

        sql = """SELECT
        cid,cname,ctype,producttype,reference,mmsi,timezone,countryofdeparture,cruisearea,startdateutc,enddateutc,vatrate,yachtname,
        yachttype,flag,imo,yachtlength,status,commercialdistancetot,commercialdistancefrcon,resultingvattot,predeliverytripnumber AS nb_pre,
        predeliverytripnumberhighseas AS nb_pre_hiseas,postdeliverytripnumber AS nb_post,postdeliverytripnumberhighseas AS nb_post_hiseas,
        commercialtripnumber AS nb_ccial,commercialtripnumberhighseas AS nb_ccial_hiseas,predeliverytripnumber + postdeliverytripnumber + commercialtripnumber AS nb_tot,
        modified,customername,contactadd_street,contactadd_street2,contactadd_zip,contactadd_city,contactadd_country,
        (commercialdistancefrcon/commercialdistancetot)*100 AS prop,commercialtripnumber - 1 AS nb_emb,commercialdistancenonfr,
        predeliverytripnumberhighseas+postdeliverytripnumberhighseas+commercialtripnumberhighseas AS nb_hiseas
        FROM "flyn_ops"."premium_ch" WHERE cid ='""" + str(contractid) + """';"""

        cursor.execute(sql)
        contract_info = list(cursor)
        a = list(contract_info)

    except:
            print('error: ', str(sys.exc_info()))
            logger.info('Error on connection to Processed Positions FLYN database.')

    fiscal_status(ctype = 'CHARTER', cert = a)
    timezone_change(timezone = a[0]['timezone'], cert = a, ctype = 'CHARTER')
    trip_calculations(cert = a)

    a[0]['rr'] = round((a[0]['prop'])*(a[0]['vatrate'])/100, 2)
    a[0]['prop'] = round(a[0]['prop'],2)

    change_name(cert = a)

    foldername =(a[0]['cname'] + "_MMSI_" + str(a[0]['mmsi']) + "_" + str(a[0]['cid']) + "_CH")
    path = 'C:\project_data\templates' + str(year) + '\__certificate_making_CH\\' + str(foldername)
    length=len(path)-len(foldername)

    folder_creation(ctype = 'CHARTER', cert = a)
    call_qgis(cert = a)
    mail_merge(ctype = 'CHARTER', cert = a)

    print("\nEND OF SCRIPT")

#------------------------------------------------------------------------------
#                               TRANSPORT CERTIFICATE
#------------------------------------------------------------------------------
elif option == '2':

    year_path(year = year, ctype = 'TRANSPORT')
    logger.info('TRANSPORT certificate initiated.')
    contractid = input('Enter yacht contract number: ')
    logger.info('Contract input: ' + str(contractid))

    try:

        sql = """SELECT DISTINCT
        c.timezone,c.name AS cname,c.id,c.reference,c.type,c.startdateutc,c.enddateutc,
        c.yachttype,c.mmsi,c.countryofdeparture,c.vatrate,c.cruisearea,ffa.status,ffa.resultinghighseasratiotext,
        ffa.predeliverytripnumber AS nb_pre,ffa.predeliverytripnumberhighseas AS nb_pre_hiseas,ffa.postdeliverytripnumber AS nb_post,
        ffa.postdeliverytripnumberhighseas AS nb_post_hiseas,ffa.commercialtripnumber AS nb_ccial,
        ffa.commercialtripnumberhighseas AS nb_ccial_hiseas,ffa.resultingvattot,vi.flag,vi.length,vi.imo,vi.name AS yachtname,
        cust.customername,cust.contactadd_street,cust.contactadd_street2,cust.contactadd_zip,cust.contactadd_city,cust.contactadd_country,
        cust.exonerated_intl, cust.commercialtripindex, ffa.commercialtripnumber - 1 AS nb_emb
        FROM "FLYNContractsData"."FLYNContracts"  AS c LEFT JOIN "FLYNContractsData"."FLYNFiscalAuto" AS ffa ON c.id = ffa.contractid
        LEFT JOIN "FLYNContractsData"."Vessel_info" AS vi ON c.mmsi = vi.mmsi
        LEFT JOIN "flyn_ops"."premium_tr3" AS cust ON c.mmsi = cust.mmsi
        WHERE c.id ='""" + str(contractid) + """' AND ffa.resultinghighseasratiotext LIKE 'High-seas trips%' AND
        cust.commercialtripindex = 1 ;"""
        cursor.execute(sql)
        b = cursor.fetchall()

        sql_cor_ex = """SELECT exonerationruletext FROM "FLYNContractsData"."FLYNFiscalAuto"
        WHERE contractid ='""" + str(contractid) + """' AND exonerationruletext LIKE
        'VAT is exempted for Continent%' OR contractid =' """ + str(contractid) + """
        ' AND exonerationruletext LIKE 'VAT is exempted for Corsica%' ;"""

        cursor.execute(sql_cor_ex)
        cor_exemption = cursor.fetchall()

    except:
            print('error: ', str(sys.exc_info()))
            logger.info('Error on connection to Processed Positions FLYN database.')

    lines = len(cor_exemption)

    if lines > 0:

        list=[]
        n = 0
        while n < lines:
            check = (cor_exemption[n]['exonerationruletext'])
            n = n+1
            list.append(check)
            myString = ""
            for elem in list:
                myString = myString + str(elem) + " "
                if 'VAT is exempted for Continent to Corsicatransportation' or 'VAT is exempted for Corsica to Continent transportation' in myString:
                    Corsica_exemption = 'YES'
                    logger.info('Corsica_exemption has been set to YES.')
    else:
        Corsica_exemption = 'NO'
        logger.info('Corsica_exemption has been set to NO.')

    if b[0]['exonerated_intl'] == 'false':
        b[0]['INTERNATIONAL_exemption'] = 'NO'
        logger.info('International exemption has been set to NO.')
    else:
        b[0]['INTERNATIONAL_exemption'] = 'YES'
        logger.info('International exemption has been set to YES.')

    b[0]['Corsica_exemption'] = Corsica_exemption

    fiscal_status(ctype = 'TRANSPORT', cert = b)
    timezone_change(timezone = b[0]['timezone'], cert = b, ctype = 'TRANSPORT')

    b[0]['nb_hiseas'] = b[0]['nb_pre_hiseas'] + b[0]['nb_post_hiseas'] + b[0]['nb_ccial_hiseas']
    b[0]['rr'] = round((b[0]['resultingvattot'])*100, 2)

    trip_calculations(cert = b)
    change_name(cert = b)

    foldername = (b[0]['cname']) + "_MMSI_" + str(b[0]['mmsi']) + "_" + str(b[0]['id']) + "_TR"
    path = 'C:\project_data\templates' + str(year) + '\__certificate_making_TR\\'+ str(foldername)
    length=len(path)-len(foldername)

    folder_creation(ctype = 'TRANSPORT', cert = b)
    call_qgis(cert = b)
    mail_merge(ctype = 'TRANSPORT', cert = b)

    print("\nEND OF SCRIPT")

#------------------------------------------------------------------------------
#                     INTERNATIONAL SALE CERTIFICATE
#------------------------------------------------------------------------------
elif option == '3':

    year_path(year = year, ctype = 'INTERNATIONAL_SALE')
    logger.info('INTERNATIONAL SALE certificate initiated.')
    contract_number = input("Please enter the contract id:\n")
    logger.info('Contract input: ' + str(contract_number))

    try:

        sql = """SELECT sale.timezone,sale.yachttype AS TYPE,sale.yachtname,sale.mmsi,sale.cname,
        sale.customername,sale.c_street,sale.c_street2,sale.c_zip,sale.c_city,sale.c_country,sale.buyername,
        sale.b_street,sale.b_street2,sale.b_zip,sale.b_city,sale.b_country,sale.sellername,sale.s_street,sale.s_street2,
        sale.s_zip,sale.s_city,sale.s_country,sale.ctype,sale.cid,sale.cruisearea,DATE(sale.startdateutc) AS "date_of_sale",
        sale.startdateutc AS "approx_start_utc",sale.enddateutc AS "approx_end_utc",vi.length,vi.flag,vi.imo,
        'EXPERT' AS "Signatory_name"
        FROM "flyn_ops"."premium_int_sale" AS sale LEFT JOIN "FLYNContractsData"."Vessel_info" AS vi ON sale.mmsi = vi.mmsi
        WHERE sale.ctype = 'INTERNATIONAL_SALE' AND sale.cid ='""" + str(contract_number) + """' ; """

        cursor.execute(sql)
        c = cursor.fetchall()
        print("\nYou are creating an INTERNATIONAL SALE certificate for: " + str(c[0]['type']) + " " + str(c[0]['yachtname']))

        sql_int = """SELECT datetime,lon,lat,nature,zone0_name FROM "FLYNPositions"."FLYNMTAutoPositions"
        WHERE  mmsi = '""" + str(c[0]['mmsi']) + """' AND nature = 'CALC' AND ZONE0_name LIKE '%int%'
        AND datetime >= '""" + str(c[0]['approx_start_utc']) + """' AND datetime <= '""" + str(c[0]['approx_end_utc']) + """'; """

        cursor.execute(sql_int)
        pos_intwat = cursor.fetchall()

        sql_table = """SELECT datetime,lon,lat,zone0_name FROM "FLYNPositions"."FLYNMTAutoPositions"
        WHERE  mmsi = '""" + str(c[0]['mmsi']) + """' AND datetime >= '""" + str(c[0]['approx_start_utc']) + """'
        AND datetime <= '""" + str(c[0]['approx_end_utc']) + """' ORDER BY datetime; """

        cursor.execute(sql_table)
        table_pos = cursor.fetchall()

        df_pos= pd.DataFrame(table_pos)
        df_pos['datetime'] = df_pos['datetime'].dt.strftime('%Y-%m-%d %H:%M')
        df_pos.drop_duplicates(subset='datetime', keep='first', inplace=True)
        df_pos.columns = ['DATE & TIME (Z) UTC', 'LATITUDE', 'LONGITUDE','LOCATION']
        df_final = df_pos.round({'LATITUDE':6, 'LONGITUDE':6})

    except:

        print('error: ', str(sys.exc_info()))

    result_returned = len(pos_intwat)

    print(f"\n{Back.MAGENTA}Search has found " + str(result_returned) + f" {Back.MAGENTA}position(s) in the database that intersect(s) international water limits:{Style.RESET_ALL} ")
    logger.info('Automatic search of DB has found '  + str(result_returned) + ' positions that intersect TW boundaries.')

    colorama_init()

    if result_returned > 1:

            display_pos = input(f"\nWould you like to list them?  Please type {Back.GREEN} Y {Style.RESET_ALL} or {Back.RED} N {Style.RESET_ALL}\n")

            if display_pos == 'Y' or display_pos == 'y':

               pos_ret=[]
               n = 0

               while n < result_returned:
                        print(f"{Back.CYAN}\nPosition:          {Style.RESET_ALL}   " + str(n))
                        string = str(pos_intwat[n]['datetime'])
                        print(f"{Back.CYAN}Date & Time (UTC): {Style.RESET_ALL}   " + (string[0:16]) + " (Z)")
                        print(f"{Back.CYAN}Longitude:         {Style.RESET_ALL}   " + str(pos_intwat[n]['lon']))
                        print(f"{Back.CYAN}Latitude:          {Style.RESET_ALL}   " + str(pos_intwat[n]['lat']))
                        n = n+1
            else:
               print("\nOK continuing with the script.\n")

    print("-"*30 + " TW crossing times and positions " + "-"*30 + "\n")

    if result_returned > 1:
        print("Entry date & time", str(pos_intwat[0]['datetime'])[0:16] , " | Longitude:", pos_intwat[0]['lon'],   "| Latitude:",pos_intwat[0]['lat'])
        print("Exit date & time ", str(pos_intwat[1]['datetime'])[0:16], " | Longitude:", pos_intwat[1]['lon'], "| Latitude:", pos_intwat[1]['lat'])
        print("\n" + "-"*30 + " TW crossing times and positions " + "-"*30)

        call_qgis(cert = c)

        confirmation = input(f"\nAre the automatically retrieved dates, times, latitudes and longitudes correct? \n\nPlease type {Back.GREEN} Y {Style.RESET_ALL} to continue with these details or {Back.RED} N {Style.RESET_ALL} to manually input the information:\n")

        if confirmation == 'Y' or confirmation == 'y':
            print("\nPositions have been converted into DECIMAL DEGREES for the certificate:")
            c[0]['EXIT_TW'] = str('YES')
            logger.info("Automatically retrieved positions have been accepted by user.")

        elif confirmation == 'N' or confirmation == 'n':

            manual_positions()

    elif result_returned == 0:
        print("Entry date & time", str(pos_intwat[0]['datetime'])[0:16] , "(Z) | Longitude:", pos_intwat[0]['lon'],   "| Latitude:",pos_intwat[0]['lat'])
        print(f"\nThere was {Back.CYAN}NO EXIT{Style.RESET_ALL} of International Waters")

        call_qgis(cert = c)

        con = input(f"\nAre you happy with the ENTRY POSITIONS, DATE and TIME (Z) - please type \n\n{Back.GREEN} Y {Style.RESET_ALL} to accept ENTRY data and manually enter EXIT data. \n{Back.RED} N {Style.RESET_ALL} to manually enter ALL data relating to (ENTRY and EXIT). \n{Back.BLUE} R {Style.RESET_ALL} to have the option of NO RE-ENTRY of TWs after sale.\n")

        if con == 'Y' or con == 'y':

            c[0]['EXIT_TW'] = str('YES')
            print("\nPlease enter the date, time (Z), longitude and latitude for the exit.")
            pos_intwat[1]['datetime'] = input(f"\nPlease enter manually the {Back.RED}DATE & TIME (Z) of EXIT{Style.RESET_ALL} out of international waters using the format: {Back.RED}YYYY-MM-DD HH:mm\n{Style.RESET_ALL}")
            pos_intwat[1]['lon'] = input(f"Please enter manually {Back.CYAN}EXIT LONGITUDE{Style.RESET_ALL} coordinates in decimal degrees e.g. 7.465641{Back.CYAN}\n{Style.RESET_ALL}")
            pos_intwat[1]['lat'] = input(f"Please enter manually the {Back.CYAN}EXIT LATITUDE{Style.RESET_ALL} coordinates in decimal degrees e.g. 43.495359{Back.CYAN}\n{Style.RESET_ALL}")
            print("\n\nYou have entered the following:")

            print("Entry time:", str(pos_intwat[0]['datetime'])[0:16], "(Z) | Longitude:", pos_intwat[0]['lon'], "| Latitude:", pos_intwat[0]['lat'])
            print("Exit time: ", str(pos_intwat[1]['datetime'])[0:16], "(Z) | Longitude:", pos_intwat[1]['lon'], "| Latitude", pos_intwat[1]['lat'])
            print("\nPositions have been converted into DECIMAL DEGREES for the certificate:\n")

            pos_intwat[1]['datetime'] = datetime.strptime(pos_intwat[1]['datetime'], '%Y-%m-%d %H:%M')
            pos_intwat[1]['lon'] = float(pos_intwat[1]['lon'])
            pos_intwat[1]['lat'] = float(pos_intwat[1]['lat'])
            logger.info("User has accepted automatically identified ENTRY date, time, lat and lon and manually entered TW EXIT information.")

        elif con == 'N' or con == 'n':
                logger.info("User has stated auto ENTRY positions aren't correct and will enter all positions, dates and times MANUALLY.")
                manual_positions()

        else:
                c[0]['EXIT_TW'] = str('NO')
                print(f"\n{Back.BLUE}User has stated no re entry of territorial waters was made after sale.{Style.RESET_ALL}")
                logger.info("User has stated no re entry of territorial waters was made after sale.")
                con2 = input(f"\nAre you happy with the ENTRY POSITIONS and TIMES (Z) - please type {Back.GREEN} Y {Style.RESET_ALL} to accept or {Back.RED} N {Style.RESET_ALL} to enter manually.\n")

                if con2 == 'y' or con2 == 'Y':
                    print("\nOK continuing without EXIT of international waters.")

                else:
                    pos_intwat[0]['datetime'] = input(f"\n\nPlease enter manually the {Back.GREEN}DATE & TIME (Z) of ENTRY{Style.RESET_ALL} into international waters using the format: {Back.GREEN}YYYY-MM-DD HH:mm\n{Style.RESET_ALL}")
                    pos_intwat[0]['lon'] = input(f"Please enter manually the {Back.CYAN}ENTRY LONGITUDE{Style.RESET_ALL} coordinates in decimal degrees e.g. 7.465641{Back.CYAN}\n{Style.RESET_ALL}")
                    pos_intwat[0]['lat'] = input(f"Please enter manually the {Back.CYAN}ENTRY LATITUDE{Style.RESET_ALL} coordinates in decimal degrees e.g. 43.495359{Back.CYAN}\n{Style.RESET_ALL}")

                    pos_intwat[0]['datetime'] = datetime.strptime(pos_intwat[0]['datetime'], '%Y-%m-%d %H:%M')
                    pos_intwat[0]['lon'] = float(pos_intwat[0]['lon'])
                    pos_intwat[0]['lat'] = float(pos_intwat[0]['lat'])

    else:
        print("\nNo positions found - all dates, times, latitudes and longitudes need to be entered manually:")

        call_qgis(cert = c)

        manual_positions()

    timezone_change(timezone = c[0]['timezone'], cert = c, ctype = 'INTERNATIONAL_SALE')

    c[0]['Enter_dtg_local'] = c[0]['Enter_dtg_local'].strftime("%Y-%m-%d %H:%M")
    c[0]['Exit_dtg_local'] = c[0]['Exit_dtg_local'].strftime("%Y-%m-%d %H:%M")
    c[0]['approx_start_local_txt'] = c[0]['approx_start_local_txt'].strftime("%Y-%m-%d %H:%M")
    c[0]['approx_end_local_txt'] = c[0]['approx_end_local_txt'].strftime("%Y-%m-%d %H:%M")

    change_name(cert = c)

    if  pos_intwat[0]['lon'] >= 0:
        c[0]['Enter_lon_txt'] = str(int(pos_intwat[0]['lon'])) + "° " + str(round((pos_intwat[0]['lon'] - int(pos_intwat[0]['lon']))*60, 3)) + "'" + "E"
    else:
        deg = str(int(pos_intwat[0]['lon']))
        deg = deg.replace('-','')
        d_min = str(round((pos_intwat[0]['lon'] - int(pos_intwat[0]['lon']))*60, 3))
        d_min = d_min.replace('-','')
        c[0]['Enter_lon_txt'] = deg + "° " + d_min + "'" + "W"

    if  pos_intwat[0]['lat'] >= 0:
        c[0]['Enter_lat_txt'] = str(int(pos_intwat[0]['lat'])) + "° " + str(round((pos_intwat[0]['lat'] - int(pos_intwat[0]['lat']))*60, 3)) + "'" + "N"
    else:
        deg = str(int(pos_intwat[0]['lat']))
        deg = deg.replace('-','')
        d_min = str(round((pos_intwat[0]['lat'] - int(pos_intwat[0]['lat']))*60, 3))
        d_min = d_min.replace('-','')
        c[0]['Enter_lat_txt'] = deg + "° " + d_min + "'" + "S"

    if  pos_intwat[1]['lon'] >= 0:
        c[0]['Exit_lon_txt'] = str(int(pos_intwat[1]['lon'])) + "° " + str(round((pos_intwat[1]['lon'] - int(pos_intwat[1]['lon']))*60, 3)) + "'" + "E"
    else:
        deg = str(int(pos_intwat[1]['lon']))
        deg = deg.replace('-','')
        d_min = str(round((pos_intwat[1]['lon'] - int(pos_intwat[1]['lon']))*60, 3))
        d_min = d_min.replace('-','')
        c[0]['Exit_lon_txt'] = deg + "° " + d_min + "'" + "W"

    if  pos_intwat[1]['lat'] >= 0:
        c[0]['Exit_lat_txt'] = str(int(pos_intwat[1]['lat'])) + "° " + str(round((pos_intwat[1]['lat'] - int(pos_intwat[1]['lat']))*60, 3)) + "'" + "N"
    else:
        deg = str(int(pos_intwat[1]['lat']))
        deg = deg.replace('-','')
        d_min = str(round((pos_intwat[1]['lat'] - int(pos_intwat[1]['lat']))*60, 3))
        d_min = d_min.replace('-','')
        c[0]['Exit_lat_txt'] = deg + "° " + d_min + "'" + "S"

    print("ENTRY lat & lon in certificate:  " + c[0]['Enter_lat_txt'] + " ; " + c[0]['Enter_lon_txt'])
    print("EXIT lat & lon in certificate:   " + c[0]['Exit_lat_txt'] + " ; " + c[0]['Exit_lon_txt'])

    now = datetime.now()
    c[0]['year'] = now.strftime("%Y")
    c[0]['Contract_number'] = c[0]['cid']
    c[0]['Cruise_area'] = c[0]['cruisearea']

    foldername = (c[0]['yachtname']) + "_MMSI_" + str(c[0]['mmsi']) + "_" + str(c[0]['Contract_number']) + "_SALE\\"
    path = 'C:\project_data\templates' + str(year) + '\__certificate_making_international_sale\\'+ str(foldername)
    length=len(path)-len(foldername)

    folder_creation(ctype = 'INTERNATIONAL_SALE', cert = c)

    mail_merge(ctype = 'INTERNATIONAL_SALE', cert = 'c')

    print("\nQGIS is loading ... for a 2nd time to allow saving of images to folder.")

    call(r"C:\Program Files\QGIS 3.22.16\bin\qgis-ltr-bin.exe")

    print("\nEND OF SCRIPT")

#------------------------------------------------------------------------------
#                             INPUT OTHER NUMBER
#------------------------------------------------------------------------------
else:
    print("\nInvalid selection please choose a number between 1-3.")
    print("\nPROGRAM HAS FINISHED")
    logger.info("Incorrect input - number doesn't correcpond to a certificate.")
    logger.info("Number between 1-3 needs to be entered.")

logger.info('PROCESSING STOPPED')
logging.shutdown()
sys.exit()
