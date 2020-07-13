#!/usr/bin/env python3
# vim: set fileencoding=utf-8 :
""" Update the Honors Master file for a new year:

    For every honor assigned in the honors.csv file, add the year in D1 to the "Years" column (G).
    Delete 'y' and 'omit' indications from Column D.
    Increment the year in D1.

"""
import csv
from googleapiclient import discovery
from google.oauth2 import service_account
import os
from parms import Parms
from pprint import pprint
import re
import gspread

def normalize(s):
    return s.lower().replace(' ','')

def getColLabel(i):
    return 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'[i-1]



parms = Parms()

# Establish Google API authorization
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

creds = service_account.Credentials.from_service_account_file(
    parms.googleapifile, scopes=SCOPES)


os.chdir(parms.datadir)

# Get all assigned honors from last year
lyhonors = set()
with open(parms.honorscsv, newline='') as csvfile:
    rdr = csv.DictReader(csvfile)
    for row in rdr:
        lyhonors.add(row['HonorID'])

# Now, update the master file.
request = discovery.build('sheets', 'v4', credentials=creds).spreadsheets().values()
sheetid = parms.honorsmaster
if '/' in sheetid:
    # Have a whole URL; get the key
    sheetid = re.search(r'/spreadsheets/d/([a-zA-Z0-9-_]+)', sheetid).groups()[0]

sheetid = '1z59UcCaDUUPa5DmEvXvFxa6nkMQ8xlTznd57XNwhb6U'

# Get the labels
labels = request.get(spreadsheetId=sheetid, range='A1:Z1').execute()['values']
labels = [normalize(l) for l in labels[0]]


# Find the year - it's right after "Alternative"
altcol = labels.index('alternative')
altcollabel = getColLabel(altcol)
year = labels[altcol+1]

# And find the 'Years' and 'Honor ID'
yearscol = labels.index('years')
yearscollabel = getColLabel(yearscol)
idcol = labels.index('honorid')

# Now go through the rest of the file and update as required

rownum = 1
while True:
    rownum += 1
    range = f'A{rownum}:Z{rownum}'
    row = request.get(spreadsheetId=sheetid, range=range).execute()
    values = row['values'][0]
    if values[idcol] in lyhonors:
        if not values[yearscol].endswith(year):
            if values[yearscol]:
                values[yearscol] += ',' + year
            else:
                values[yearscol] = year
        values[altcol] = ''
        request.update(spreadsheetId=sheetid, range=range, body=values).execute()
        pprint(response)

