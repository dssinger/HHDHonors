#!/usr/bin/env python
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
import textwrap

infile = open('honors.csv', 'r')
batfile = open('sendem.sh', 'w')
batfile.write('#!/bin/sh\n')
rdr = csv.DictReader(infile)
linenum = 0
for line in rdr:
    linenum += 1
    outhtml = []
    outhtml.append('Dear %s,' % line['Dear'])
    parts = ["L'shanah Tovah!  Thank you for participating in our High Holy Day services with the honor"]
    parts.append('<b>%s</b> at <b>%s</b> Services on <b>%s</b> at <b>%s</b>.' % (line['Honor'], line['Service'], line['Service_Date'], line['Location']))
    if line['Sharing']:
        parts.append('You will be joined in this honor by %s.' % line['Sharing'])
    outhtml.append(' '.join(parts))
    parts = ["Please confirm that you've received this note"]
    if line['Filename']:
        parts.append("and the attached cue sheet")
    parts.append('by replying to <a href="mailto:honors@shirhadash.org?subject=Confirming%20%receipt%20of%20honors%20notification">honors@shirhadash.org</a>.')
    parts.append('If you have questions about this honor or cannot accept it, please contact <a href="mailto:rabbiaron@shirhadash.org?subject=Questions%20about%20High%20Holy%20Day%20Honor">Rabbi Aron</a>.')
    outhtml.append(' '.join(parts))
    parts = ['On %s, please arrive at %s (%s minutes before the beginning of the worship service) in order to meet with %s, who will review your participation in the service and answer questions about seating and cues.' % (line['Holiday'], line['Arrive'], line['Early'], line['Rabbi'])]
    if line['Filename']:
        parts.append('Please remember to bring the attached cue sheet with you to the service.')
    outhtml.append(' '.join(parts))
    
    if line['FromText']:
        outhtml.append('Your reading is in <b>%s</b>, <b>%s</b>.  Read the English from "<b>%s</b>" to "<b>%s</b>".' % (line['Book'], line['Pages'], line['FromText'], line['ToText'])) 

    outhtml.append('We anticipate an inspiring and meaningful High Holy Day season.  Thank you for joining with us to add to the special flavor of our services at Congregation Shir Hadash.')
            
    

    outfn = '%03d' % linenum
    outf = open(outfn+'.html', 'w')
    outf.write('<html>\n<body>\n<p>')
    outf.write('</p>\n<p>'.join(outhtml))
    outf.write('</p>\n</body>\n</html>\n')
    outf.close()
    outf = open(outfn+'.txt', 'w')
    
    outtext = [textwrap.fill(re.sub(r'<.*?>', '', x),72) for x in outhtml]

    outf.write('\n\n'.join(outtext))
    outf.close()

    outtext = './sendmail.py --to %s %s ' % (line['Email1'], line['Email2'])
    outtext += ' --htmlfile ../%s.html --textfile ../%s.txt ' % (outfn, outfn)
    if line['Filename']:
        outtext += ' --attach %s ' % line['Filename']
    batfile.write(outtext)
    batfile.write('\n')
    batfile.write('sleep 10\n')

batfile.close()
