#!/usr/bin/env python3
# vim: set fileencoding=utf-8 :
""" Make the files to be used for email """
import xlrd
import sys
import csv
import datetime
import re
import os
import textwrap
from urllib.parse import quote_plus, quote

mformurl = "https://docs.google.com/forms/d/e/1FAIpQLSfw_P9yYEWCLHNraTnlTLvChAUfci6l3Dd7y_1F-4U4bqWqMA/viewform?usp=pp_url&entry.1032808708=honortodisplay&entry.365808493=name&entry.172610770=service&entry.1586648944=honorid&entry.177841022=honordesc"


if  __name__ == '__main__':
    from parms import Parms
    parms = Parms()
    os.chdir(parms.datadir)
    
    sourcedir = os.path.dirname(os.path.abspath(__file__))
    
    infile = open(parms.honorscsv, 'r')
    try:
        os.mkdir(parms.maildir)
    except FileExistsError:
        pass
    os.chdir(parms.maildir)
    
    batfile = open('sendem.sh', 'w')
    batfile.write('#!/bin/sh\n')
    batfile.write('cd "%s"\n' % os.path.join(parms.datadir, parms.maildir))
    rdr = csv.DictReader(infile)
    linenum = 0
    for line in rdr:
        linenum += 1
        outhtml = []
        formurl = mformurl
        formurl = formurl.replace("honortodisplay",quote_plus("%s (%s)" % (line['Honor'], line['Service'])))
        formurl = formurl.replace("name", quote_plus(line['Full_Name']))
        formurl = formurl.replace("honorid", line['HonorID'])
        formurl = formurl.replace("honordesc", quote_plus(line['Honor']))
        formurl = formurl.replace("service", quote_plus(line['Service']))
        emailsub = quote('HHD Honor ' + line['HonorID'] + ': ' + line['Honor'] + ' (' + line['Service'] + ')')
        outhtml.append('<p>Dear %s,</p>' % line['Dear'])
        parts = ["<p>L'shanah Tovah!  It is my great pleasure to invite you to participate in our High Holy Day services with the honor"]
        parts.append('<b>%s</b> at <b>%s</b> Services on <b>%s</b> at <b>%s</b>.' % (line['Honor'], line['Service'], line['Service_Date'], line['Location']))
        if line['Sharing']:
            parts.append('You will be joined in this honor by %s.' % line['Sharing'])
        parts.append('</p>')
        parts.append('<p>I\'m pleased that we are able to recognize your special contribution to the life of our congregation in this way and express our appreciation for your dedication in the past year.</p>')
        outhtml.append(' '.join(parts))
        parts = []
        parts.append('<div id="onlyhtml">')
        parts.append("Please let us know if you will be able to accept this honor by:")
        parts.append('<ul>')
        parts.append('<li>Responding at <a href="%s">this link</a>, or</li>' % formurl)
        parts.append('<li>Emailing <a href="mailto:honors@shirhadash.org?subject=' + emailsub + '">honors@shirhadash.org</a>, or</li>')
        parts.append('<li>Calling Gloria Clouss at the Temple Office, 408-358-1751 x7.</li>')
        parts.append('</ul>')
        parts.append('</div>')
        parts.append('<p>If you have questions about this honor, please contact <a href="mailto:rabbiaron@shirhadash.org?subject=Questions%20about%20High%20Holy%20Day%20Honor%20' + emailsub + '">Rabbi Aron</a>.</p>')
        outhtml.append(' '.join(parts))
        outhtml.append('<p>On %s, please arrive at %s (%s minutes before the beginning of the worship service) in order to meet with %s, who will review your participation in the service and answer questions about seating and cues.</p>' % (line['Holiday'], line['Arrive'], line['Early'], line['Rabbi']))
        if line['Filename']:
            outhtml.append('<p>Please remember to print the attached cue sheet and bring it with you to the service.</p>')
        if line['Explanation']:
            outhtml.append('<p>%s</p>' % line['Explanation'])
        if line['FromText']:
            outhtml.append('<p>Your reading is in <b>%s</b>, <b>%s</b>.  Read the English from "<b>%s</b>" to "<b>%s</b>".</p>' % (line['Book'], line['Pages'], line['FromText'], line['ToText'])) 
        outhtml.append('<p>We anticipate an inspiring and meaningful High Holy Day season.  Thank you for joining with us to add to the special flavor of our services at Congregation Shir Hadash.</p>')

        # Prepare the plain-text alternative.
        
        textonly = 'Please call Gloria Clouss at the Temple Office, 408-358-1751 x7, to let us know if you will be able to accept this honor.'
        text = [re.sub(r'<div id="onlyhtml">.*?</div>', textonly, x) for x in outhtml]
        
        text = [re.sub(r'<br.*?>', '\n\n', x) for x in text]
    
        outtext = '\n\n'.join([textwrap.fill(re.sub(r'<.*?>', '', x),72) for x in text])
        
        # Add signatures to both alternatives.

        outhtml.append('<p>\n<p><p>Sincerely,</p><p>\n</p><p>Naomi Parker<br />President</p>')
        outtext += '\n\nSincerely,\n\nNaomi Parker, President'
    

        outfn = '%03d' % linenum
        outf = open(outfn+'.html', 'w')
        outf.write('<html>\n<body>\n<p>')
        outf.write('</p>\n<p>'.join(outhtml))
        outf.write('</p>\n</body>\n</html>\n')
        outf.close()
        
        
        outf = open(outfn+'.txt', 'w')
        outf.write(outtext)
        outf.close()

        outtext = '"%s" --YMLfile "%s" --to %s %s ' % (os.path.join(sourcedir, 'sendmail.py'), os.path.join(sourcedir, 'cshmail.yml'), line['Email1'], line['Email2'])
        outtext += ' --htmlfile ./%s.html --textfile ./%s.txt ' % (outfn, outfn)
        if line['Filename']:
            outtext += ' --attach "%s" ' % os.path.join(parms.cuedir, line['Filename'])
        batfile.write(outtext)
        batfile.write('\n')
        batfile.write('sleep 10\n')

    batfile.close()
