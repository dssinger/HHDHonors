#!/usr/bin/env python
import os, sys, xlrd
from people import Nickname, People, getLabelsFromSheet, stringify
from generatehonors import Honor

from apiclient.discovery import build
from httplib2 import Http
from oauth2client import file, client, tools
from gsheet import GSheet
from pprint import pprint

apikey = 'AIzaSyCvKjNGOxDw7KgAWyc_VjGdw7tlx7i4x84'
honorsfile = '1pbffRKiaG7lMVeSii9EFLKay4KVMU2pmPQVozAOs9tM'

Nickname.setnicknames()
people = People.loadpeople('/Users/david/Dropbox/High Holy Day Honors/2018/2018 Roster for Honors.xlsx')

def makeLabelsDict(row):
    """Returns all of the labels from a spreadsheet as a dict
     Labels are normalized by converting them to lower case,
        removing any leading "home-", and
        removing all spaces. """
    labels = []
    for p in row:
        p = p.lower()
        if p.startswith('home-'):
            p = p[5:]
        p = ''.join(p.split())
        labels.append(p)  
    ret = dict(zip(labels, xrange(len(labels))))
    # Provide two-way associativity
    for p in ret.keys():
      ret[ret[p]] = p
    return ret

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
        