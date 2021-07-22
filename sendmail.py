#!/usr/bin/env python3
""" Send an email contained in a file. """
import cshparse, os, sys, argparse, smtplib, time
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart 
from email.mime.application import MIMEApplication


def avoiddups(which):
  res = []
  for item in which:
    cur.execute(f"SELECT COUNT(*) FROM mailed WHERE sentto = ? AND htmlfile = ? AND textfile = ?", (item, parms.htmlfile, parms.textfile))
    count = cur.fetchone()[0]
    if count == 0:
      res.append(item)
    elif parms.verbose >= 2:
      print(f"not duplicating {item} for {parms.htmlfile} ({parms.textfile})")
  return res  

import collections.abc
def flatten(l):
    ### From http://stackoverflow.com/questions/2158395/flatten-an-irregular-list-of-lists-in-python
    for el in l:
        if isinstance(el, collections.abc.Iterable) and not isinstance(el, str):
            for sub in flatten(el):
                yield sub
        else:
            yield el


# Handle parameters
parms = cshparse.cshparse(description=__doc__, YMLfile="cshmail.yml", includedbparms=False)
parms.parser.add_argument("--htmlfile", dest='htmlfile', default='')
parms.parser.add_argument("--textfile", dest='textfile', default='')
parms.parser.add_argument("--mailserver", dest='mailserver')
parms.parser.add_argument("--mailpw", dest='mailpw')
parms.parser.add_argument("--mailport", dest='mailport')
parms.parser.add_argument("--from", dest='from')
parms.parser.add_argument("--to", dest='to', nargs='+', default=[], action='append')
parms.parser.add_argument("--cc", dest='cc', nargs='+', default=[], action='append')
parms.parser.add_argument("--bcc", dest='bcc', nargs='+', default=[], action='append')
parms.parser.add_argument("--subject", dest='subject', default='Shir Hadash High Holy Day Honor for You')
parms.parser.add_argument("--attachment", dest='attachment', nargs='+', default=[], action='append')
parms.parser.add_argument("--dry-run", dest='dryrun', action='store_true')
parms.parser.add_argument("--sleep", dest='sleep', type=float, default=3)
parms.parser.add_argument("--verbose", "-v", dest='verbose', default=1, action='count')
parms.parser.add_argument("--quiet", "-q", dest='quiet', default=0, action='count')

parms.parse()
parms.sender = parms.__dict__['from']  # Get around reserved word

parms.verbose = parms.verbose - parms.quiet  # Resolve verbosity level

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

# Remove targets who have already gotten this mail:
import sqlite3
conn = sqlite3.connect('mail.db')
cur = conn.cursor()
cur.execute('CREATE TABLE IF NOT EXISTS mailed (sentto text, htmlfile text, textfile text)')

parms.to = avoiddups(parms.to)
parms.cc = avoiddups(parms.cc)
parms.bcc = avoiddups(parms.bcc)




# And send the mail.
targets = []
if parms.to:
    targets.extend(parms.to)
if parms.cc:
    targets.extend(parms.cc)
if parms.bcc:
    targets.extend(parms.bcc)



if targets and not parms.dryrun:
    # Connect to the mail server:
    mailconn = smtplib.SMTP_SSL(parms.mailserver, parms.mailport)
    mailconn.ehlo()
    mailconn.login(parms.sender, parms.mailpw)

    # and send the mail
    mailconn.sendmail(parms.sender, targets, finalmsg)

    # and sleep
    time.sleep(parms.sleep)

# If all went well, log the transaction
res = []
for which in (parms.to, parms.cc, parms.bcc):
  for item in which:
    res.append((item, parms.htmlfile, parms.textfile))
    if parms.verbose >= 1:
      print(f'{"Would send" if parms.dryrun else "Sent"} {parms.htmlfile} {parms.textfile} to {item}')

if res:
  cur.executemany('INSERT INTO mailed VALUES (?, ?, ?) ', res)

conn.commit()
