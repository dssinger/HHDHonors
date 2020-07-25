#!/usr/bin/env python3
# vim: set fileencoding=utf-8 :
""" Make the files to be used for email

    This is the special 2020 version for Zoom services.  Things that are different:
    * If two people are sharing an honor, the first one reads the prayers before the reading and the second
      reads the prayers after the reading.
    * Except honor 6320, where the first person reads Hebrew and the second reads English.
    * We do not ask people to call the office.
    """

import xlrd
import sys
import csv
import datetime
import re
import os
import textwrap
from urllib.parse import quote_plus, quote



if  __name__ == '__main__':
    from parms import Parms
    parms = Parms()
    mformurl = f"https://docs.google.com/forms/d/e/{parms.responseform}/viewform?usp=pp_url&entry.1032808708=honortodisplay&entry.365808493=name&entry.172610770=service&entry.1586648944=honorid&entry.177841022=honordesc"
    
    sourcedir = os.path.dirname(os.path.abspath(__file__))
    print(sourcedir)
    print(os.path.abspath(__file__))
    
    os.chdir(parms.datadir)
    infile = open(parms.honorscsv, 'r')
    try:
        os.mkdir(parms.maildir)
    except FileExistsError:
        pass
    os.chdir(parms.maildir)
    
    batfile = open('sendem.sh', 'w')
    os.chmod('sendem.sh', 0o755)
    batfile.write('#!/bin/sh\n')
    batfile.write('cd "%s"\n' % os.path.join(parms.datadir, parms.maildir))
    rdr = csv.DictReader(infile)
    linenum = 0
    sharer = 0
    lasthonor = None
    for line in rdr:
        linenum += 1
        if line['HonorID'] == lasthonor:
            sharer += 1
        else:
            lasthonor = line['HonorID']
            sharer = 0
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
        parts.append('<b>%s</b> at <b>%s</b> Services on <b>%s</b>.' % (line['Honor'], line['Service'], line['Service_Date']))
        if line['Sharing'] and line['HonorID'] < '8000':
            shabbatnote = ''
            if line['HonorID'] == '6320':
                options = ('Hebrew text', 'English text')
            else:
                if line['Honor'].startswith("Haftarah"):
                    scroll = 'Haftarah'
                    if line['Filename'].split('.')[0].endswith('s') and sharer == 1:
                        shabbatnote = "Please note the special wording for Shabbat in the blessing after the Haftarah reading.  Contact the Cantor if you would like a recording of this blessing."
                else:
                    scroll = 'Torah'
                options = (f'blessing before the {scroll} is read', f'blessing after the {scroll} is read')

            parts.append(f'You will read the {options[sharer]}.  ({line["Sharing"]} will read the {options[::-1][sharer]}.)')
            if shabbatnote:
                parts.append(f'</p>\n{shabbatnote}')
        parts.append('</p>')
        parts.append('<p>I\'m pleased that we are able to recognize your special contribution to the life of our congregation in this way and express our appreciation for your dedication in the past year.</p>')
        outhtml.append(' '.join(parts))
        parts = []
        parts.append('<div id="onlyhtml">')
        parts.append("Please let us know if you will be able to accept this honor by:")
        parts.append('<ul>')
        parts.append('<li>Responding at <a href="%s">this link</a>, or</li>' % formurl)
        parts.append('<li>Emailing <a href="mailto:honors@shirhadash.org?subject=' + emailsub + '">honors@shirhadash.org</a></li>')
        parts.append('</ul>')
        parts.append('</div>')
        parts.append(f'<p>If you have questions about this honor, please contact <a href="mailto:{parms.rabbiemail}?subject=Questions%20about%20High%20Holy%20Day%20Honor%20{emailsub}">{parms.rabbiname}</a>.</p>')
        outhtml.append('\n'.join(parts))
        outhtml.append(f'<p><b>Before</b> {line["Service"]}, please go to the <a href="https://shirhadash.org/">Shir Hadash website</a> and sign up for the service so you will receive the Zoom link in advance.  ')
        outhtml.append(f'<b>On</b> {line["Service"]}, please come online at {line["Arrive"]} to practice unmuting and be ready for your participation.')
        if line['Filename']:
            outhtml.append('<p>Please remember to print the attached cue sheet and have it handy during the Zoom call.</p>')
        if line['Explanation']:
            outhtml.append('<p>%s</p>' % line['Explanation'])
        if line['FromText']:
            outhtml.append('<p>Your reading is in <b>%s</b>, <b>%s</b>.  Read the English from "<b>%s</b>" to "<b>%s</b>".</p>' % (line['Book'], line['Pages'], line['FromText'], line['ToText'])) 
        outhtml.append('<p>We anticipate an inspiring and meaningful High Holy Day season.  Thank you for joining with us to add to the special flavor of our services at Congregation Shir Hadash.</p>')

        # Prepare the plain-text alternative.
        
        textonly = f'Please email honors@shirhadash.org to let us know if you will be able to accept this honor.'
        text = [re.sub(r'<div id="onlyhtml">.*?</div>', textonly, x) for x in outhtml]
        
        text = [re.sub(r'<br.*?>', '\n\n', x) for x in text]

        outtext = '\n\n'.join([textwrap.fill(re.sub(r'<.*?>', '', x),72) for x in text])
        
        # Add signatures to both alternatives.

        outhtml.append(f'<p>\n<p><p>Sincerely,</p><p>\n</p><p>{parms.president}<br />President</p>')
        outtext += f'\n\nSincerely,\n\n{parms.president}, President'
    
        if line['Filename']:
            outfn = line['Filename'].split('.')[0]
        else:
            outfn = line['HonorID']
        if line['Sharing']:
            outfn += '-' + str(sharer+1)
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
