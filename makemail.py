#!/usr/bin/env python3
# vim: set fileencoding=utf-8 :
""" Make the files to be used for email

    This is the special 2020 version for Zoom services.  Things that are different:
    * If two people are sharing an honor, the first one reads the prayers before the reading and the second
      reads the prayers after the reading.
    * Except honor 6320, where the first person reads Hebrew and the second reads English.
    * We do not ask people to call the office.
    """


import csv
import re
import os
import textwrap
import sys
from urllib.parse import quote_plus, quote


if __name__ == '__main__':
    from parms import Parms
    parms = Parms()
    parms.onlydo = sys.argv[1:]

    mformurl = f'https://docs.google.com/forms/d/e/{parms.responseform}/viewform?usp=pp_url&' \
               f'entry.1032808708=honortodisplay&' \
               f'entry.365808493=name&' \
               f'entry.172610770=service&' \
               f'entry.1586648944=honorid&' \
               f'entry.177841022=honordesc'
    
    sourcedir = os.path.dirname(os.path.abspath(__file__))
    
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
    prevreader = ''
    sleep = False
    for line in rdr:
        linenum += 1
        fullhonor = line['HonorID'].split('-')[0]
        if fullhonor == lasthonor:
            sharer += 1
        else:
            lasthonor = fullhonor
            sharer = 0
        outhtml = []
        formurl = mformurl
        formurl = formurl.replace("honortodisplay", quote_plus(f"{line['Honor']} ({line['Service']})"))
        formurl = formurl.replace("name", quote_plus(line['Full_Name']))
        formurl = formurl.replace("honorid", line['HonorID'])
        formurl = formurl.replace("honordesc", quote_plus(line['Honor']))
        formurl = formurl.replace("service", quote_plus(line['Service']))
        emailsub = quote('HHD Honor ' + line['HonorID'] + ': ' + line['Honor'] + ' (' + line['Service'] + ')')

        cueexp = ''
        subhonor = ''
        shabbatnote = ''
        if line['Sharing']:
            if line['HonorID'] < '8240':
                if line['HonorID'] == '6320':
                    subhonor = f" ({('Hebrew text', 'English text')[sharer]})"
                else:
                    if line['Honor'].startswith("Haftarah"):
                        scroll = 'Haftarah'
                        if line['Filename'].split('.')[0].split('-')[0].endswith('s') and sharer == 1:
                            shabbatnote = "<p>Please note the special wording for Shabbat in the blessing after the " \
                                          "Haftarah reading.  Contact the Cantor if you would like a recording of " \
                                          "this blessing.</p>"
                    else:
                        scroll = 'Torah'
                    subhonor = f" (Blessing {('before', 'after')[sharer]} the {scroll} reading)"
            else:
                subhonor = f" (Part {sharer+1})"


        outhtml.append(f'<p>Dear {line["Dear"]},</p>')
        outhtml.append(
            f"<p>L'shanah Tovah!  It is my great pleasure to invite you to participate in our High Holy Day services"
            f" on Zoom with the honor"
            f" <b>{line['Honor']}{subhonor}</b> at <b>{line['Service']} Services</b>"
            f" on <b>{line['Service_Date']}</b>.</p>")
        if shabbatnote:
            outhtml.append(f"<p>{shabbatnote}</p>")
        if line['HonorID'] >= '8240':
            cueexp = line['Cue'] if line['Cue'] else prevreader
        prevreader = line['Full_Name']

        outhtml.append('<p>I\'m pleased that we are able to recognize your special contribution to the life of our '
                     'congregation in this way and express our appreciation for your dedication in the past year.</p>')
        outhtml.append('\n'.join(
                 ('<div id="onlyhtml">', "Please let us know if you will be able to accept this honor by:",
                 '<ul>',
                 f'<li>Responding at <a href="{formurl}">this link</a>, or</li>',
                 f'<li>Emailing <a href="mailto:honors@shirhadash.org?subject={emailsub}">honors@shirhadash.org</a></li>'
                 '</ul>',
                 '</div>')))
        outhtml.append(
                 f'<p>If you have questions about this honor, please contact '
                 f'<a href="mailto:{parms.rabbiemail}'
                 f'?subject=Questions%20about%20High%20Holy%20Day%20Honor%20{emailsub}">'
                 f'{parms.rabbiname}</a>.</p>')
        outhtml.append(f'<p><b>Before</b> {line["Service"]}, '
                       f'please go to the <a href="https://shirhadash.org/">Shir Hadash website</a> '
                       f'and sign up for the service so you will receive the Zoom link in advance.  ')
        outhtml.append(f'<b>On</b> {line["Service"]}, please come online at {line["Arrive"]} '
                       f'to practice unmuting and be ready for your participation.')
        if line['Filename']:
            outhtml.append('<p>Please remember to print the attached cue sheet '
                           'and have it handy during the service on Zoom.</p>')
            if cueexp:
                outhtml.append(f'<p>You will read after {cueexp}.</p>')
        if line['Explanation']:
            outhtml.append(f'<p>{line["Explanation"]}</p>')
        if line['FromText']:
            outhtml.append(
                f'<p>Your reading is in <b>{line["Book"]}</b>, <b>{line["Pages"]}</b>.  '
                f'Read the English from "<b>{line["FromText"]}</b>" to "<b>{line["ToText"]}</b>".</p>')
        outhtml.append('<p>We anticipate an inspiring and meaningful High Holy Day season.  '
                       'Thank you for joining with us to add to the special flavor of our services at '
                       'Congregation Shir Hadash.</p>')

        # Prepare the plain-text alternative.
        
        textonly = 'Please email honors@shirhadash.org to let us know if you will be able to accept this honor.'
        text = [re.sub(r'(?s)<div id="onlyhtml">.*?</div>', textonly, x) for x in outhtml]
        
        text = [re.sub(r'<br.*?>', '\n\n', x) for x in text]

        outtext = '\n\n'.join([textwrap.fill(re.sub(r'<.*?>', '', x), 72) for x in text])
        
        # Add signatures to both alternatives.

        outhtml.append(f'<p>\n<p><p>Sincerely,</p><p>\n</p><p>{parms.president}<br />President</p>')
        outtext += f'\n\nSincerely,\n\n{parms.president}, President'

        outfn = line['HonorID']
        if line['Sharing'] and '-' not in outfn:
            outfn = f'{outfn}-{sharer+1}'
        outf = open(outfn+'.html', 'w')
        outf.write('<html>\n<body>\n<p>')
        outf.write('</p>\n<p>'.join(outhtml))
        outf.write('</p>\n</body>\n</html>\n')
        outf.close()
        
        outf = open(outfn+'.txt', 'w')
        outf.write(outtext)
        outf.close()

        if parms.onlydo and outfn not in parms.onlydo:
            continue

        if sleep:
            batfile.write('sleep 3\n')
        sleep = True   

        outtext = f'"{os.path.join(sourcedir, "sendmail.py")}" ' \
                  f'--YMLfile "{os.path.join(sourcedir, "cshmail.yml")}" ' \
                  f'--to {line["Email1"]} {line["Email2"]} ' \
                  f' --htmlfile ./{outfn}.html ' \
                  f'--textfile ./{outfn}.txt '
        if line['Filename']:
            outtext += f' --attach "{os.path.join(parms.cuedir, line["Filename"])}" '
        batfile.write(outtext)
        batfile.write('\n')

    batfile.close()
