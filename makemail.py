#!/usr/bin/env python3
# vim: set fileencoding=utf-8 :
""" Make the files to be used for email
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

    mformurl = f'{parms.responseform}/viewform?usp=pp_url&' \
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
    batfile.write('#!/bin/bash\n')
    batfile.write('cd "%s"\n' % os.path.join(parms.datadir, parms.maildir))
    batlist = []
    rdr = csv.DictReader(infile)
    linenum = 0
    sharer = 0
    lasthonor = None
    prevreader = ''
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
                elif line['Honor'].startswith("Haftarah"):
                    scroll = 'Haftarah'
                    if line['Filename'].split('.')[0].split('-')[0].endswith('s') and sharer == 1:
                        shabbatnote = "<p>Please note the special wording for Shabbat in the blessing after the " \
                                      "Haftarah reading.  Contact the Cantor if you would like a recording of " \
                                      "this blessing.</p>"
                    subhonor = f" (Blessing {('before', 'after')[sharer]} the {scroll} reading)"
                elif line['Honor'].startswith("Torah"):
                    scroll = 'Torah'
                    subhonor = f" (Blessing {('before', 'after')[sharer]} the {scroll} reading)"
            else:
                subhonor = f" (Part {sharer+1})"


        outhtml.append(f'<p>Dear {line["Dear"]},</p>')
        outhtml.append(
            f"<p>L'shanah Tovah!  It is my great pleasure to invite you to participate in our High Holy Day services"
            f" at Shir Hadash"
            f" with the honor"
            f" <b>{line['Honor']}{subhonor}</b> at the <b>{line['Service']} Service</b>"
            f" on <b>{line['Service_Date']}</b>.")
        if line['Sharing'] and line['Filename']:
            outhtml.append(f"  You will be sharing this honor with {line['Sharing']}.")
        outhtml.append('</p>')
        if shabbatnote:
            outhtml.append(f"<p>{shabbatnote}</p>")
        if line['HonorID'] >= '8140' and line['HonorID'] < '9000':
            cueexp = line['Cue'] if line['Cue'] else prevreader
        prevreader = line['Full_Name']

        outhtml.append('<p>I\'m pleased that we are able to recognize your special contribution to the life of our '
                     'congregation in this way and express our appreciation for your dedication in the past year.</p>')
        outhtml.append(f'<p>All honors will be conducted in person this year; we cannot accommodate remote honors.')
        outhtml.append(f'A seat will be reserved in the Sanctuary for you in the Honorees Section for the {line["Service"]} Service.  '
                       f'We request that you sit in your designated seat prior to and <b>after</b> your honor. '
                       f'Sanctuary seating will be limited and those with Honors at a service are being given the '
                       f'opportunity to have their immediate family members present in the Sanctuary for that service. '
                       f'Other seating will be available in the Chapel, Oneg Room and outdoors.</p>')
        outhtml.append('\n'.join(
                 ('<div id="onlyhtml">', "Please let us know if you will be able to accept this honor by:",
                 '<ul>',
                 f'<li>Responding at <a href="{formurl}">this link</a>, or</li>',
                 f'<li>Emailing <a href="mailto:honors@shirhadash.org?subject={emailsub}">honors@shirhadash.org</a> (please include the number of seats you would like in the Sanctuary for your <b>immediate</b> family)</li>'
                 '</ul>',
                 '</div>')))
        outhtml.append(
                 f'<p>If you have questions about this honor, please contact '
                 f'<a href="mailto:{parms.rabbiemail}'
                 f'?subject=Questions%20about%20High%20Holy%20Day%20Honor%20{emailsub}">'
                 f'{parms.rabbiname}</a>.</p>')
        outhtml.append(f"<p>On {line['Holiday']}, please arrive at {line['Arrive']} ({line['Early']} minutes before"
                       f" the beginning of the worship service) in order to meet with {line['Rabbi']},"
                       f" who will review your participation in the service"
                       f" and answer questions about seating and cues.</p>")
        if line['Filename']:
            outhtml.append('<p>Please remember to print the attached cue sheet '
                           'and bring it with you to the service.</p>')
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
        
        textonly = 'Please email honors@shirhadash.org to let us know if you will be able to accept this honor'
        textonly += '  (please include the number of seats to reserve in the Sanctuary for your immediate family).'
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

        if parms.onlydo:
            print(f"{outfn}: {line['Full_Name']} - {line['Email1']} {line['Email2']}")


        outtext = f'"{os.path.join(sourcedir, "sendmail.py")}" ' \
                  f'--YMLfile "{os.path.join(sourcedir, "cshmail.yml")}" ' \
                  f'--to {line["Email1"]} {line["Email2"]} ' \
                  f' --htmlfile ./{outfn}.html ' \
                  f'--textfile ./{outfn}.txt '
        if line['Filename']:
            fn = os.path.join(parms.cuedir, line["Filename"])
            if os.path.exists(fn):
                outtext += f' --attach "{fn}"'
            else:
                print(f'Could not find {fn} or {line["Filename"]} in {parms.cuedir}')
        batlist.append(outtext)

    batlist[-1] += ' --sleep 0\n'  # The last entry doesn't need sleep
    batfile.write('\n'.join(batlist))
    batfile.close()
