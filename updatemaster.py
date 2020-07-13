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
    return str(s).lower().replace(' ','')

parms = Parms()

# Need to do this before we move to the data directory:
gc = gspread.auth.service_account(filename=parms.googleapifile)



os.chdir(parms.datadir)

# Get all assigned honors from last year
lyhonors = set()
with open(parms.honorscsv, newline='') as csvfile:
    rdr = csv.DictReader(csvfile)
    for row in rdr:
        lyhonors.add(row['HonorID'])

# Now, update the master file.
sheetid = parms.honorsmaster
if '/' in sheetid:
    # Have a whole URL; get the key
    sheetid = re.search(r'/spreadsheets/d/([a-zA-Z0-9-_]+)', sheetid).groups()[0]

sheetid = '1z59UcCaDUUPa5DmEvXvFxa6nkMQ8xlTznd57XNwhb6U'

sheet = gc.open_by_key(sheetid).sheet1

# Get everything

allvalues = sheet.get_all_values(value_render_option='UNFORMATTED_VALUE')

# Get the labels
labels = [normalize(l) for l in allvalues[0]]


# Find the year - it's right after "Alternative"
altcol = labels.index('alternative')
yearcol = altcol + 1
year = str(labels[yearcol])

# And find the 'Years' and 'Honor ID'
yearscol = labels.index('years')
idcol = labels.index('honorid')

# Now go through the rest of the file and update as required

updates = []

rownum = 1
for row in allvalues[1:]:
    rownum += 1
    id = str(row[idcol])
    if id in lyhonors or (f'id{row[altcol]}' in lyhonors):
        if not row[yearscol].endswith(year):
            if row[yearscol].strip():
                row[yearscol] += ',' + year
            else:
                row[yearscol] = year
    row[yearcol] = ''
    #updates.append({'range':f'A{rownum}:Z{rownum}', 'values':[row]})

# Update the year
allvalues[0][yearcol] = int(year) + 1
# Now, do the updates

range = gspread.utils.rowcol_to_a1(1+len(allvalues), 1+len(allvalues[0]))
sheet.batch_update([{'range': f'A1:{range}','values':allvalues}])


