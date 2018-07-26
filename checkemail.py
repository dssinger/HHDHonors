#!/usr/bin/env python
import os, sys, xlrd
from people import People
from generatehonors import Honor

from apiclient.discovery import build
from httplib2 import Http
from oauth2client import file, client, tools
from gsheet import GSheet
from pprint import pprint

import yaml

parms = yaml.load(open('honors.yaml','r'))

people = People.loadpeople(parms['roster'])



# Connect to the spreadsheet and get the values

sheet = GSheet(parms['honors'], parms['apikey'])

for row in sheet:
    for name in (row.name1, row.name2):
        name = ' '.join(name.strip().split())
        if not name:
            continue
        if name=='XXX':
            continue
        if 'anniversary' in name.lower():
            continue
        if 'confirmation' in name.lower():
            continue
        if 'presidents' in name.lower():
            continue
        person = People.findbyname(name)
        if person and '@' not in person.email:
            print name, 'has no email'
        