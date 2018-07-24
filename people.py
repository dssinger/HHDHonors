#!/usr/bin/python
"""People (for HHD, based on Shul Suite data)

Creates an array of people indexed by "firstname lastname", containing their address fields.
If any names are duplicated, complains.

"""

import xlrd
import sys
from datetime import datetime, date


def getLabelsFromSheet(sheet):
    """Returns all of the labels from a spreadsheet as a dict
     Labels are normalized by converting them to lower case,
        removing any leading "home-", and
        removing all spaces. """
    labels = []
    for p in sheet.row_values(0):
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

def stringify(value):
    """ Convert values to strings """
    # Let's normalize everything to strings/unicode strings
    if isinstance(value, (int, long, float)):
        value = '%d' % value
    if isinstance(value, bool):
        value = '1' if value else '0'
    elif isinstance(value, (datetime, date)):
        value = ('%s' % value)[0:10]

    return value
    
class Nickname:
    nicknames = {}
    def __init__(self, first, last, newfirst='', newlast=''):
        self.first = first
        self.last = last
        self.newfirst = newfirst if newfirst else first
        self.newlast = newlast if newlast else last
        self.key = self.first + ' ' + self.last
        self.newkey = self.newfirst + ' ' + self.newlast
        self.nicknames[self.key] = self
        
    @classmethod
    def find(self, key):
        try:
            return self.nicknames[key]
        except KeyError:
            return None

        
    @classmethod
    def setnicknames(self):
        Nickname('Patricia', 'Uro-May', 'Pat')
        
    def __repr__(self):
        return('key: %s, first: %s => %s, last: %s => %s' % (self.key, self.first, self.newfirst, self.last, self.newlast))


class People:
    linenumber = 1
    people = {}
    debug = False
    synonyms = {
                'dr': 'Drive',
                'st': 'Street',
                'ave': 'Avenue',
                'rd': 'Road',
                'ct': 'Court',
                'blvd': 'Boulevard',
                'av': 'Avenue',
                'ln': 'Lane',
                'wy': 'Way',
                'cir': 'Circle',
                'n': 'North',
                'no': 'North',
                'PO': 'P.O.',
                'pl': 'Place',
                's': 'South',
                'w': 'West',
                'e': 'East'
            }
    @classmethod
    def setlabels(self, labels):
        self.labels = labels
        
    @classmethod
    def getlinenumber(self):
        self.linenumber += 1
        return self.linenumber
    
    @classmethod
    def loadpeople(self, fn, debug=False):
        self.debug = debug
        s = xlrd.open_workbook(fn)
        sheet = s.sheets()[0]
        self.setlabels(getLabelsFromSheet(sheet))
        for r in range(sheet.nrows-1):
            People(sheet.row_values(r+1))
            
    @classmethod
    def find(self, id):
        try:
            return self.people[id]
        except KeyError:
            print "Could not find %s" % id
            return False
            
    @classmethod
    def findbyname(self, key):
        try:
            if len(self.people[key]) > 1:
                print "%s has multiple entries; use ID!" % key
            return self.people[key][0]
        except KeyError:
            print "Could not find %s" % key
            return False
            
    def getaddr(self):
        return "%s, %s, %s  %s" % (self.streetaddress, self.city, self.state, self.postalcode)
        
    def handlenickname(self):
        """Process names where the roster has a formal name but we want to use an informal name"""
        update = Nickname.find(self.key)
        if update:
            dparts = self.displayname.split()
            for i in range(len(dparts)):
                if dparts[i] == self.firstname:
                    dparts[i] = update.newfirst
                if dparts[i] == self.lastname:
                    dparts[i] = update.newlast
            self.displayname = ' '.join(dparts)
            (self.key, self.firstname, self.lastname) = (update.newkey, update.newfirst, update.newlast)


            
	
    def __init__(self, row):
        for x in range(len(row)):
            self.__dict__[self.labels[x]] = ' '.join(stringify(row[x]).split())
        # Let's try to normalize the street address, at least for things like Dr/Dr./Drive
        
        sa = []
        for w in self.streetaddress.split():
            wc = w.replace('.','')
            wc = wc.replace(',','')
            wc = wc.lower()
            if wc in self.synonyms:
                sa.append(self.synonyms[wc] + (',' if ',' in w else ''))
            else:
                sa.append(w)
        self.streetaddress = ' '.join(sa)
        
        self.key = self.firstname + ' ' + self.lastname
        self.handlenickname()
        
        self.internalcontactid = self.getlinenumber()
        self.people[self.internalcontactid] = self
        
        
        if self.debug and self.key in self.people:
            print "Duplicate: %s" % self.key
            other = self.people[self.key]
            for one in other:
                print "  Old: ID=%5s, address=%s" % (one.internalcontactid, one.getaddr())
            print "  New: ID=%5s, address=%s" % (self.internalcontactid, self.getaddr())
            self.people[self.key].append(self)
        else:
            self.people[self.key] = [self]
 
    def __repr__(self):
        return "display = '%s', first = '%s', last = %s, address = %s, email = %s" % (self.displayname, self.firstname, self.lastname, self.streetaddress, self.email)
        

if __name__ == "__main__":
    if len(sys.argv) == 1:
        filename = 'People.xlsx'
    else:
        filename = sys.argv[1]
    Nickname.setnicknames()
    People.loadpeople(filename, debug=True)


        
		
		