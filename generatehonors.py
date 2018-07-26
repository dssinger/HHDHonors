#!/usr/bin/env python3
# vim: set fileencoding=utf-8 :
""" Create the honors.csv file for the mail merge.
  For each honor (unique to service, of course), find all those sharing
  the honor.  If they share an address, combine them.
  
  Use the Honor ID to make matches.
  
  Generate a new spreadsheet with one line per person/couple per honor,
  with a new field for "sharer".

"""
import xlrd
import sys
import csv
import datetime
import re
import os

datemode = 1  # make it global




def space(s):
  return ' '.join(s.split())


def normalizeService(service):
  """ Convert a service into full form, removing any sequence numbers and 'Service" """
  if '-' in service:
    service = service.split('-',1)[1]
  service = service.strip()
  service = service.replace('YK','Yom Kippur')
  service = service.replace('RH','Rosh Hashanah')
  service = service.replace(' Service', '')
  return service


def normalizeHonor(honor):
  """Take the full description (or short description) of an honor and:
  1)  make it lower-case
  2)  remove all non-alphamerics (including spaces)
  """
  return ''.join(re.split(r'\W+', honor)).lower()

def stringify(value):
  """ Convert values to strings """
  # Let's normalize everything to strings/unicode strings
  if isinstance(value, (int, float)):
      value = '%d' % value
  if isinstance(value, bool):
      value = '1' if value else '0'
  elif isinstance(value, (datetime.datetime, datetime.date)):
      value = ('%s' % value)[0:10]
  
  return value


class Service:
  services = {}
  labels = {}
  
  @classmethod
  def setlabels(self, labels):
    self.labels = labels
  
  @classmethod
  def find(self, key):
    return self.services[normalizeService(key)]
  
  def __init__(self, row):
    # Convert the items in the row into attributes of the object
    for x in range(len(row)):
      attrname = self.labels[x]
      if attrname not in ('time', 'arrive', 'date', 'early'):
          row[x] = stringify(row[x])
      self.__dict__[attrname] = row[x]
    
    # Normalize the service name
    self.service = normalizeService(self.service)
    self.daypart = normalizeService(self.daypart)
    
    # And make "early" an integer
    self.early = int(self.early)
    
    # Now, let's get the times and dates into pretty form and figure out if this service is on Shabbat or not.
    t = datetime.time(*xlrd.xldate_as_tuple(self.time, datemode)[3:])
    a = datetime.time(*xlrd.xldate_as_tuple(self.arrive, datemode)[3:])
    d = datetime.date(*xlrd.xldate_as_tuple(self.date, datemode)[0:3])
    
    
    if d.weekday() == 4 and t.hour >= 19:
      self.isShabbat = True
    elif d.weekday() == 5 and t.hour <= 19:
      self.isShabbat = True
    else:
      self.isShabbat = False
    
    self.time = t.strftime('%I:%M %p').lstrip('0')
    self.arrive = a.strftime('%I:%M %p').lstrip('0')
    self.date = d.strftime('%A, %B %d, %Y').replace(' 0',' ')

    
    
    
    self.services[self.service] = self

class Honor:
  honors = {}
  labels = {}
  all = []
  @classmethod
  def register(self, honoree):
    key = honoree.honor
    if key not in self.honors:
      self.honors[key] = Honor(honoree.honor, honoree.service)
    self.honors[key].sharers.append(honoree)
    self.honors[key]
    return self.honors[key]
  
  
  @classmethod
  def setlabels(self, labels):
    self.labels = labels
  
  
  
  @classmethod
  def find(self, honorid):
    return self.honors.get(honorid, None)
  
  def assign(self, person):
    keys = ['streetaddress', 'city', 'state', 'postalcode']
    address = ' '.join(person.__dict__[k] for k in keys)
    try:
      self.sharers[address].addname(person)
    except KeyError:
      self.sharers[address] = Honoree(person)

  
  
  
  def __init__(self, row):
    # This method defines an honor and puts it in the class indexed by the HonorID
    # the service has the sequence number removed, and RH and YK expanded
    #print(row)
    row.service = normalizeService(row.service)
    
    # Convert the items in the row into attributes of the object.  Make some substitutios in attribute name.
    for x in self.labels:
      attrname = x
      if attrname == 'from':
          attrname = 'fromtext'
      elif attrname == 'to':
          attrname = 'totext'
      self.__dict__[attrname] = stringify(row.__dict__[x])
    self.sharers = {}  # Indexed by address, of course
    self.honors[self.honorid] = self
    
    # See if there's a file associated with the service; if so, do Shabbat adjustment if need be.
    if self.cuesheet.lower() in ('y', 's', 'o'):
        if Service.find(self.service).isShabbat and self.cuesheet.lower() == 's':
            self.filename = self.honorid + 's.pdf'
        else:
            self.filename = self.honorid + '.pdf'
    else:
        self.filename = ''
    

    
    
    # And finally, append to the list of honors in order of appearance
    self.all.append(self)
  
  
  def __repr__(self):
    ans = "{'sharers': '%s'," % self.sharers
    ans += ', '.join(["'%s': '%s'" % (x, self.__dict__[x]) for x in list(self.labels.keys())])
    return ans

  
  
  def sharing(self):
    """ Return a list, with a dict in each position:
      me: me
      them: a comma separated list of other sharers' names, with 'and' at the end
    """
    res = []
    sharers = [self.sharers[k] for k in self.sharers]
    sharers.sort(key=lambda x: x.dear)
    for i in range(len(sharers)):
      entry = {}
      entry['me']= sharers[i]
      them = [sharers[j].fullnames for j in range(len(sharers)) if j != i]
      if len(them) == 0:
        entry['them'] = ''
      elif len(them) == 1:
        entry['them'] = them[0]
      elif len(them) == 2:
        entry['them'] = ' and '.join(them)
      else:
        them[-1] = 'and ' + them[-1]
        entry['them'] = ', '.join(them)
      res.append(entry)
    return res



class Honoree:
  def __init__(self, person):
    self.firstname = person.firstname.strip()
    self.lastname = person.lastname.strip()
    self.fullname = person.displayname
    self.fullnames = self.fullname
    self.addr = person.streetaddress
    self.city = person.city
    self.state = person.state
    self.zip = person.postalcode
    self.csz = self.city + ' ' + self.state + ', ' + self.zip
    self.address = '\n'.join([self.addr, self.csz])
    self.dear = self.firstname
    self.sendpaper = person.sendpaper
    try:
        self.email1 = person.email
    except AttributeError:
        self.email1 = ''
    self.email2 = ''
  
  def addname(self, person):
    self.firstname = [self.firstname, person.firstname]
    self.dear = ' and '.join(self.firstname)
    if self.lastname != person.lastname:
      self.lastname = [self.lastname, person.lastname]
      self.fullname = [self.fullname, person.firstname + ' ' + person.lastname]
      self.fullnames = ' and '.join(self.fullname)
    else:
      self.fullnames = ' and '.join(self.firstname) + ' ' + self.lastname
    try:
      self.email2 = person.email
    except AttributeError:
      self.email2 = ''
    
    
 


def getLabelsFromSheet(sheet):
  """Returns all of the labels from a spreadsheet as a dict"""
  labels = [''.join(p.split()).lower() for p in sheet.row_values(0)]
  ret = dict(list(zip(labels, list(range(len(labels))))))
  # Provide two-way associativity
  for p in list(ret.keys()):
    ret[ret[p]] = p
  return ret
  
if __name__ == '__main__':
    
    import os
    from gsheet import GSheet
    from parms import Parms
    parms = Parms()
    os.chdir(parms.datadir)

    # We need to load service information first, because that determines whether we use Shabbat files or regular ones.
    # The "Services Master.xls" file has information about each service, which we use in preference to that in the HHD Honors file.
    services = xlrd.open_workbook(parms.services)
    datemode = services.datemode
    sheet = services.sheets()[0]
    # Put the labels into the Service class
    Service.setlabels(getLabelsFromSheet(sheet))

    # Now, load the services into the class
    for r in range(sheet.nrows-1):
      Service(sheet.row_values(r+1))
      
    # Now, load membership data
    from people import People
    People.loadpeople(parms.roster)

    # The Honors file now has honors AND assignmenbts, and it's a Google spreadsheet instead of an Excel file.
    
    master = GSheet(parms.honors, parms.apikey)
    Honor.setlabels(master.labels)
    
    
    for row in master:
        # Have we already defined this honor?
        honor = Honor.find(row.honorid)
        if not honor:
            honor = Honor(row)
            
        # Do the assignments
        for name in (row.name1, row.name2):
            if name:
                name = ' '.join(name.strip().split())
                lname = name.lower()
                if lname != 'xxx' and 'anniversary' not in lname and 'confirmation' not in lname:
                    person = People.findbyname(name)
                    if person:
                        honor.assign(person)
                    




    ## OK, now we can create the updated spreadsheet
    ##   and build the batch file to print the cue sheets

    outfile = csv.writer(open(parms.honorscsv, 'w'))
    outfile.writerow(('Dear',
      'Full_Name',
      'Family_Address',
      'Family_CSZ',
      'Email1',
      'Email2',
      'Send_Paper',
      'Service',
      'Service_Date',
      'Service_Time',
      'Location',
      'Early',
      'Arrive',
      'Holiday',
      'Rabbi',
      'HonorID',
      'Honor',
      'Book',
      'Pages',
      'Cue',
      'FromText',
      'ToText',
      'Filename',
      'Sharing',
      'Explanation'))
  
    cuesheets = {}
    gscmd = "gs -sDEVICE=pdfwrite -dPDFSETTINGS=/default -dNOPAUSE -dQUIET -dBATCH -sOutputFile=-"
    for theHonor in Honor.all:
      if theHonor.honorforletter.lower() == 'none':
          continue  # Skip honors that don't need letters generated.
     
      if not theHonor.sharing():
          continue  # Skip honors with no assignees
  
      # Figure out page and book.
      if theHonor.pagestart:
          theHonor.book = 'Mishkan Hanefesh'
          if theHonor.pageend:
              pages = 'pages %s-%s' % (theHonor.pagestart, theHonor.pageend)
          else:
              pages = 'page %s' % (theHonor.pagestart)
      else:
          theHonor.book = ''
          pages = ''
  
      theService = Service.services[theHonor.service]
  
      honorname = theHonor.honorforletter.strip()
      if not honorname:
          honorname = theHonor.description
  
      for s in theHonor.sharing():
          if '@' not in s['me'].email1  or (s['me'].email2 and '@' not in s['me'].email2):
              print('Missing email for %s' % s['me'].fullnames)
              s['me'].sendpaper = True 
          
          outfile.writerow((s['me'].dear,
          s['me'].fullnames,
          s['me'].addr,
          s['me'].csz,
          s['me'].email1,
          s['me'].email2,
          s['me'].sendpaper,
          theService.service,
          theService.date,
          theService.time,
          theService.location,
          theService.early,
          theService.arrive,
          theService.daypart,
          theService.rabbi,
          theHonor.honorid,
          honorname.replace('\\u2026','...'),
          theHonor.book,
          pages,
          theHonor.cue.replace('\\u2026','...'),
          theHonor.fromtext,
          theHonor.totext,
          theHonor.filename,
          s['them'],
          theHonor.explanation.replace('\\u2026','...')))
      
          if theHonor.filename:
              snum = theHonor.filename[0]
              if snum not in cuesheets:
                  cuesheets[snum] = []
              cuesheets[snum].append((theHonor.filename, honorname, s['me'].fullnames))
          
          
    # Finally, write the printing batch files
    import codecs

    for k in cuesheets:
        batfile = codecs.open('print%s.sh' % k,'w', encoding='utf-8')
        batfile.write('#!/bin/bash\n')
        batfile.write('cd "../New HHD Cues and Readings"\n')
        batfile.write('gs -dBATCH -dNOPAUSE -q -sDEVICE=pdfwrite -sOutputFile=%s.pdf \\\n' % k)
        #cuesheets[k].reverse()
        for s in cuesheets[k]:
            batfile.write('    %s\\\n' % s[0])

        batfile.write('\n')
        for s in cuesheets[k]:
            batfile.write('# %s - %s - %s\n' % s)
        batfile.close()
    

    
