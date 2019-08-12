#!/bin/sh
mkdir "$HOME/Dropbox/HHD Cues/partsfootered"
cd "$HOME/Dropbox/HHD Cues/parts"


for i in 2040*.pdf
  do
    gs -o ./footer.pdf  -sDEVICE=pdfwrite -c "/Helvetica findfont 14 scalefont setfont" -c "50 50 moveto ($i) show showpage"
    pdftk "$i" stamp footer.pdf output "../partsfootered/$i"
  done

