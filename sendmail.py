#!/usr/bin/env python
""" Send an email contained in a file. """
import tmparms, os, sys, argparse, smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart 
from email.mime.application import MIMEApplication

import tmglobals
globals = tmglobals.tmglobals()

import collections
def flatten(l):
    ### From http://stackoverflow.com/questions/2158395/flatten-an-irregular-list-of-lists-in-python
    for el in l:
        if isinstance(el, collections.Iterable) and not isinstance(el, basestring):
            for sub in flatten(el):
                yield sub
        else:
            yield el


# Handle parameters
parms = tmparms.tmparms(description=__doc__, YMLfile="tmmail.yml", includedbparms=False)
parms.parser.add_argument("--htmlfile", dest='htmlfile')
parms.parser.add_argument("--textfile", dest='textfile')
parms.parser.add_argument("--mailserver", dest='mailserver')
parms.parser.add_argument("--mailpw", dest='mailpw')
parms.parser.add_argument("--mailport", dest='mailport')
parms.parser.add_argument("--from", dest='from')
parms.parser.add_argument("--to", dest='to', nargs='+', default=[], action='append')
parms.parser.add_argument("--cc", dest='cc', nargs='+', default=[], action='append')
parms.parser.add_argument("--bcc", dest='bcc', nargs='+', default=[], action='append')
parms.parser.add_argument("--subject", dest='subject', default='Shir Hadash High Holiday Information')
parms.parser.add_argument("--attachment", dest='attachment', nargs='+', default=[], action='append')

globals.setup(parms, connect=False, gotodatadir=False)

parms.sender = parms.__dict__['from']  # Get around reserved word

if parms.attachment:
    # Create wrapper and main message part
    msg = MIMEMultipart('mixed')
    alternatives = MIMEMultipart('alternative')
else:
    # Just create main message part
    msg = MIMEMultipart('alternative')
    alternatives = msg


msg['Subject'] = parms.subject
msg['From'] = parms.sender

# Flatten recipient lists and insert to and cc into the message header
parms.to = list(flatten(parms.to))
parms.cc = list(flatten(parms.cc))
parms.bcc = list(flatten(parms.bcc))
parms.attachment = list(flatten(parms.attachment))
msg['To'] = ', '.join(parms.to)
msg['cc'] = ', '.join(parms.cc)




# Now, create the alternative parts.
if parms.textfile:
    part1 = MIMEText(open(parms.textfile, 'r').read(), 'plain')
    alternatives.attach(part1)
else:
    alternatives.attach(MIMEText('This is a multipart message with no plain-text part', 'plain'))
    
if parms.htmlfile:
    part2 = MIMEText(open(parms.htmlfile, 'r').read(), 'html')
    alternatives.attach(part2)
    
if parms.attachment:
    msg.attach(alternatives)
    for item in parms.attachment:
        # PDFs only for now
        stuff = open(item, 'rb').read()
        data = MIMEApplication(stuff, 'pdf')
        data.add_header('Content-Disposition', 'attachment', filename = os.path.basename(item))
        msg.attach(data)


# Convert the message to string format:
finalmsg = msg.as_string()

# And send the mail.
targets = []
if parms.to:
    targets.extend(parms.to)
if parms.cc:
    targets.extend(parms.cc)
if parms.bcc:
    targets.extend(parms.bcc)



# Connect to the mail server:
mailconn = smtplib.SMTP_SSL(parms.mailserver, parms.mailport)
mailconn.ehlo()
mailconn.login(parms.sender, parms.mailpw)

# and send the mail
mailconn.sendmail(parms.sender, targets, finalmsg)


