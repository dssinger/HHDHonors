#!/usr/bin/env python3
"""People (for HHD, based on ShulCloud data)

Creates an array of people indexed by "firstname lastname", containing their address fields.
If any names are duplicated, complains.

ShulCloud has ONE record per household, containing all adults in the household.

"""

from openpyxl import load_workbook
import re
from datetime import datetime, date



def getLabelsFromSheet(sheet):
    """Returns all of the labels from a spreadsheet as a dict
     Labels are normalized by converting them to lower case,
        removing any leading "home-",
        removing any "'s"
        replacing any '_' with spaces,
        removing any non-alphamerics (including spaces),
        replacing spaces with '_'.
        We treat 'first_name', 'last_name', and 'id' specially to avoid problems elsewhere in the code.
        """
    labels = []
    for p in [cell.value for cell in sheet[1]]:
        p = p.lower()
        if p.startswith('home-'):
            p = p[5:]
        p = p.replace("'s", "")
        p = p.replace('_', ' ')
        p = '_'.join(re.split(r'\W+', p))
        if p.startswith('_'):
            p = p[1:]
        if p.endswith('_'):
            p = p[:-1]
        p = p.replace('first_name', 'firstname')
        p = p.replace('last_name', 'lastname')
        if p == 'id':
            p = 'household_id'
        labels.append(p)  
    ret = dict(list(zip(labels, list(range(len(labels))))))
    # Provide two-way associativity
    for p in list(ret.keys()):
      ret[ret[p]] = p
    return ret

def stringify(value):
    """ Convert values to strings """
    # Let's normalize everything to strings/unicode strings
    if isinstance(value, (int, float)):
        value = '%d' % value
    if isinstance(value, bool):
        value = '1' if value else '0'
    elif isinstance(value, (datetime, date)):
        value = ('%s' % value)[0:10]

    return str(value).strip()

def normalize(value):
    # Convert to string, and get rid of extraneous spaces
    return ' '.join(stringify(value).split())
    
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
        return
        Nickname('Ivan (Rusty)', 'Gralnik', 'Rusty')
        Nickname('S. Henry', 'Stern', 'Henry')
        Nickname('Itzhak', 'Nir', 'Itzik')
        Nickname('Isabelle', 'Schneider', 'Billee')
        Nickname('Rebecca Katz', 'Tedesco', 'Rebecca', 'Katz Tedesco')
        Nickname('Hugh', 'Seid-Valencia', 'Rabbi Hugh', 'Seid-Valencia')
        Nickname('Julia', 'Hartman', 'Julie')
        Nickname('Marilyn', 'Keelan', 'Lindy')
        Nickname('Michael', 'Adelman', 'Mickey')
        Nickname('Jeffrey', 'Segol', 'Jeff')
        Nickname('Adrian E', 'Cerda', 'Adrian')
        Nickname('Stephen', 'Jackson', 'Steve')
        
        
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
            
    namesyns = (('Bob', 'Robert', 'Robbie', 'Rob'),
                ('Margie', 'Marjorie'),
                ('Joe', 'Joseph'),
                ('Irv', 'Irving'),
                ('Andy', 'Andrew'),
                ('Rich', 'Richard'),
                ('Josh', 'Joshua'),
                ('Cindy', 'Cynthia'),
                ('Kim', 'Kimberly'),
                ('Ed', 'Edward'),
                ('Nick', 'Nicholas'),
                ('Debbie', 'Deborah', 'Deb', 'Debra'),
                ('Zach', 'Zack', 'Zachary'),
                ('Rochelle', 'Shell'),
                ('Barb', 'Barbara', 'Barbra'),
                ('Liz', 'Elizabeth', 'Elisabeth'),
                ('Mickey', 'Mike', 'Michael'),
                ('John', 'Jon', 'Jonathan'),
                ('Mady', 'Madeleine'),
                ('Pat', 'Patricia')
                )
    namesyns = ()
                
    Nickname.setnicknames()
                
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
        commonfields = ('household_id', 'address', 'address2', 'city', 'state', 'zip' )
        personfields = ('email', 'title', 'firstname', 'lastname', 'nickname')
        firstnamecol = personfields.index('firstname')
        nicknamecol = personfields.index('nickname')
        emailcol = personfields.index('email')
        thefields = personfields + commonfields
        p1fields = ['primary_' + f for f in personfields] + list(commonfields)
        p2fields = ['secondary_' + f for f in personfields] + list(commonfields)
        s = load_workbook(fn)
        sheet = s.active
        self.setlabels(getLabelsFromSheet(sheet))
        for inrow in sheet.iter_rows(min_row=2):
            # Clean up the values and put them in a dictionary indexed by column label
            values = [normalize(v.value) for v in inrow]
            row = {self.labels[i]: values[i] for i in range(len(values))}

            # if only the primary in a row has an email, give it to the secondary, too
            if not row['secondary_email']:
                row['secondary_email'] = row['primary_email']

            for fgroup in (p1fields, p2fields):
                # Load each person, taking firstname synonyms into account
                altnames = False
                # Handle firstname synonyms:
                pinfo = [row[field] for field in fgroup]
                People(pinfo, thefields)
                if  pinfo[nicknamecol] and pinfo[nicknamecol] != pinfo[firstnamecol]:
                    People(pinfo, thefields, firstname = pinfo[nicknamecol])
            
    @classmethod
    def find(self, id):
        try:
            return self.people[id]
        except KeyError:
            print("Could not find %s" % id)
            return False
            
    @classmethod
    def findbyname(self, key):        
        try:
            if len(self.people[key]) > 1:
                print("%s has multiple entries; use ID!" % key)
            return self.people[key][0]
        except KeyError:
            # print("Could not find %s" % key)
            return False
            
    def getaddr(self):
        return "%s, %s, %s  %s" % (self.streetaddress, self.city, self.state, self.postalcode)
        
    def handlenickname(self):
        return
        """Process names where the roster has a formal name but we want to use an informal name"""
        if self.nickname and self.firstname != self.nickname:
            self.displayname = self.nickname + ' ' + self.lastname




	
    def __init__(self, row, labels, firstname=None, lastname=None):
        for x in range(len(row)):
            setattr(self, labels[x], row[x])

        # Handle special processing for names if needed:
        if firstname:
            self.firstname = firstname
        if lastname:
            self.lastname = lastname

        # Let's try to normalize the street address, at least for things like Dr/Dr./Drive
        sa = []
        for w in self.address.split():
            wc = w.replace('.','')
            wc = wc.replace(',','')
            wc = wc.lower()
            if wc in self.synonyms:
                sa.append(self.synonyms[wc] + (',' if ',' in w else ''))
            else:
                sa.append(w)
        self.address = ' '.join(sa)
        
        self.key = self.firstname + ' ' + self.lastname
        self.displayname = self.key
        self.handlenickname()
        
        self.internalcontactid = self.getlinenumber()
        self.people[self.internalcontactid] = self
    
        
        if self.debug and self.key in self.people:
            if not firstname and not lastname:
                print("Duplicate: %s" % self.key)
                other = self.people[self.key]
                for one in other:
                    print("  Old: ID=%5s, address=%s" % (one.internalcontactid, one.getaddr()))
                print("  New: ID=%5s, address=%s" % (self.internalcontactid, self.getaddr()))
            self.people[self.key].append(self)
        else:
            self.people[self.key] = [self]
 
    def __rexpr__(self):
        return "display = '%s', first = '%s', last = '%s', address = '%s', email = '%s', sendpaper = %s" % (self.displayname, self.firstname, self.lastname, self.streetaddress, self.email, self.sendpaper)
        

if __name__ == "__main__":
    from parms import Parms
    import os
    parms = Parms()
    os.chdir(parms.datadir)
    People.loadpeople(parms.roster)
    for person in People.people:
        print(person)



        
		
		
