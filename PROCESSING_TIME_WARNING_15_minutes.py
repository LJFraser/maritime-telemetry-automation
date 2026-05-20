# -*- coding: utf-8 -*-
"""
Created on Mon Dec 11 12:25:49 2023

@author: Leane Fraser
"""
#MODULES REQUIRED
import psycopg2
import psycopg2.extras
import win32com.client as win32
from datetime import timedelta, datetime
import logging

#GENERATES THE LOG FILE - lists only the times when the processing time is greater than 5 minutes
logging.basicConfig(filename='C:\project_data\templates\processing_checks.log', encoding='utf-8', level=logging.DEBUG)
logger = logging.getLogger()
logger.setLevel(logging.DEBUG)

#CONNECTION to  DATABASE

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
cursor = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)

#SQL QUERY
sql = """SELECT nb_processed_contracts,last_process_duration FROM "database"."branche_of_database" """
cursor.execute(sql)
processing = cursor.fetchall()

now = datetime.now()
d = processing[0]['last_process_duration']
no_con = processing[0]['nb_processed_contracts']

if d > timedelta(minutes=5): #The limit for the processing time is set here
    logger.info(f'{now} -- PROCESSING TIME greater than 5 minutes - last_process_duration: {d}')

#IF THE PROCESSING TIME IS > 5 minutes the email is launched
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = 'operations-team1@example.com' #production tests
#   mail.To = 'operations-team1@example.com;operations-team2@example.com;operations-team3@example.com'     #activate when building .exe
    mail.Subject = f'PROCESSING TIME IS GREATER THAN 15 minutes ({d}) WARNING - automatic email'

#HTML tags are used in the email for formatting
    mail.HTMLBody += f"""<div>
                            <h1 style="font-family: 'Calibri'; font-size: 23; font-weight: bold; color: #0A506E;">PROCESSING TIME WARNING > 15 minutes.</h1>
                        </div>
                        <div><i><h9 style="font-family: 'Calibri'; font-size: 13; color: #129090;">Number of contracts processing: {no_con} </h9></i></div>
                        <div><i><h9 style="font-family: 'Calibri'; font-size: 13; color: #129090;">Processing time: {d} </h9></i></div>"""
    mail.Send()
