#!/usr/bin/env python
import os, sys, xlrd
from people import People
from generatehonors import Honor

from apiclient.discovery import build
from httplib2 import Http
from oauth2client import file, client, tools
from gsheet import GSheet
from pprint import pprint

apikey = 'AIzaSyCvKjNGOxDw7KgAWyc_VjGdw7tlx7i4x84'
honorsfile = '1pbffRKiaG7lMVeSii9EFLKay4KVMU2pmPQVozAOs9tM'

people = People.loadpeople('/Users/david/Dropbox/High Holy Day Honors/2018/2018 Roster for Honors.xlsx')



# Connect to the spreadsheet and get the values

sheet = GSheet(honorsfile, apikey)

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
        