#!/bin/sh
cd "$HOME/Dropbox/HHD Cues/nofooters/"


for i in *.pdf
  do
    gs -o ./footer.pdf  -sDEVICE=pdfwrite -c "/Helvetica findfont 14 scalefont setfont" -c "50 20 moveto ($i) show showpage"
    pdftk "$i" stamp footer.pdf output "../parts/$i"
  done
rm footer.pdf ../parts/footer.pdf

