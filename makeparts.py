#!/usr/bin/env python3

import xlrd, sys, os
mazchor = '/Users/david/Dropbox/HHD Cues/Mishkan HaNefesh'
rhpages = os.path.join(mazchor, 'Rosh HaShanah')
ykpages = os.path.join(mazchor, 'Yom Kippur')
workbook = '/Users/david/Dropbox/High Holy Day Honors/src/HHD Honors Master.xlsx'
joincmd= '"/System/Library/Automator/Combine PDF Pages.action/Contents/Resources/join.py"'

outdir = os.path.normpath(os.path.join(mazchor,'..','parts'))
oldparts = '/Users/david/Dropbox/High Holy Day Honors/New HHD Cues and Readings'
book = xlrd.open_workbook(workbook)
sheet = book.sheet_by_index(0)

# Map the column names
def normalize(s):
    return s.lower().replace(' ','')
    
colnames = [normalize(item) for item in sheet.row_values(0)]    
    
class Honor:
    def __init__(self, row):
        for i in range(len(colnames)):
            try:
                if row[i].ctype == 2:
                    self.__dict__[colnames[i]] = int(row[i].value)
                elif row[i].ctype == 4:
                    self.__dict__[colnames[i]] = (row[i].value == 1)
                else:
                    self.__dict__[colnames[i]] = row[i].value
            except IndexError:
                self.__dict__[colnames[i]] = ''
                
                
        # Clean up cuesheet indication
        self.cuesheet = normalize(self.cuesheet)
        
        # Identify mazchor
        if self.honorid < 2000:
            self.mazchor = ''
        elif self.honorid < 5000:
            self.mazchor = rhpages
        else:
            self.mazchor = ykpages
        
        out = []
        if self.cuesheet:
            if self.pagestart:
                if isinstance(self.pageend, int):
                    pages = [self.pagestart, self.pageend]
                    input = ' '.join(['"%s/%0.3d.pdf"' % (self.mazchor, page) for page in pages])
                    out.append('%s -o "%s/%s.pdf" %s' % (joincmd, outdir, self.honorid, input))
                    if self.cuesheet == 's':
                        out.append('%s -o "%s/%ss.pdf" %s' % (joincmd, outdir, self.honorid, input))
                else:
                    out.append('cp "%s/%0.3d.pdf" "%s/%s.pdf"' % (self.mazchor, self.pagestart, outdir, self.honorid))
                    if self.cuesheet == 's':
                        out.append('cp "%s/%0.3d.pdf" "%s/%ss.pdf"' % (self.mazchor, self.pagestart, outdir, self.honorid))
            else:
                out.append('cp "%s/%s.pdf" "%s"' % (oldparts, self.honorid, outdir))
                if self.cuesheet == 's':
                    out.append('cp "%s/%ss.pdf" "%s"' % (oldparts, self.honorid, outdir))
                    
        if out:
            self.cmd = '\n'.join(out) + '\n'
        else:
            self.cmd = ''
                    
                
    def __repr__(self):
        return ', '.join(['%s = %s' % (item, h.__dict__[item]) for item in colnames])
        
                
                
outfile = open('doit.sh', 'w')

for row in sheet.get_rows():
    if row[0].ctype == 2:
        h = Honor(row)
        outfile.write(h.cmd)
       
      
        