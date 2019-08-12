#!/usr/bin/env python3

import sys, os, gsheet
from parms import Parms
mazchor = '/Users/david/Dropbox/HHD Cues/Mishkan HaNefesh'
rhpages = os.path.join(mazchor, 'Rosh HaShanah')
ykpages = os.path.join(mazchor, 'Yom Kippur')
workbook = '/Users/david/Dropbox/High Holy Day Honors/src/HHD Honors Master.xlsx'
joincmd= '"/System/Library/Automator/Combine PDF Pages.action/Contents/Resources/join.py"'

outdir = os.path.normpath(os.path.join(mazchor,'..','parts'))
oldparts = '/Users/david/Dropbox/High Holy Day Honors/New HHD Cues and Readings'

parms = Parms()

sheet = gsheet.GSheet(parms.HHDMaster, parms.apikey)

def makenum(s):
    if s.strip():
        return int(s)
    else:
        return s

def buildcmd(row):

    honorid = makenum(row.honorid)
    row.pagestart = makenum(row.pagestart)
    row.pageend = makenum(row.pageend)
    row.honorid = row.honorid + row.alternative.strip()

    # Identify mazchor
    if honorid < 2000:
        row.mazchor = ''
    elif honorid < 5000:
        row.mazchor = rhpages
    else:
        row.mazchor = ykpages
    
    out = []
    if row.cuesheet:
        if row.pagestart:
            if isinstance(row.pageend, int) and (row.pagestart != row.pageend):
                pages = [row.pagestart, row.pageend]
                input = ' '.join(['"%s/%0.3d.pdf"' % (row.mazchor, page) for page in pages])
                out.append('%s -o "%s/%s.pdf" %s' % (joincmd, outdir, row.honorid, input))
                if row.cuesheet == 's':
                    out.append('%s -o "%s/%ss.pdf" %s' % (joincmd, outdir, row.honorid, input))
            else:
                out.append('cp "%s/%0.3d.pdf" "%s/%s.pdf"' % (row.mazchor, row.pagestart, outdir, row.honorid))
                if row.cuesheet == 's':
                    out.append('cp "%s/%0.3d.pdf" "%s/%ss.pdf"' % (row.mazchor, row.pagestart, outdir, row.honorid))
        else:
            out.append('cp "%s/%s.pdf" "%s/"' % (oldparts, row.honorid, outdir))
            if row.cuesheet == 's':
                out.append('cp "%s/%ss.pdf" "%s/"' % (oldparts, row.honorid, outdir))
                
    if out:
        return '\n'.join(out) + '\n'
    else:
        return ''
                

                
                
outfile = open('makeallparts.sh', 'w')
outfile.write('mkdir "%s"\n' % outdir)

for row in sheet:
    c = buildcmd(row)
    if c:
        if 'cp' not in c:
            outfile.write('echo %s\n' %c[1:-1])
        outfile.write(c)


       
      
        