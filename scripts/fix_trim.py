"""Trim every page of a PDF to exactly 6 x 9 in.

LibreOffice snaps a 9 in page height to 649.16 pt (229 mm). The extra
1.16 pt sits below the bottom margin, so it is cropped off there.
"""
import sys
from pypdf import PdfReader, PdfWriter

W, H = 432, 648
src = sys.argv[1]
reader = PdfReader(src)
writer = PdfWriter(clone_from=reader)
for page in writer.pages:
    box = page.mediabox
    top = float(box.top)
    for b in (page.mediabox, page.cropbox, page.trimbox, page.bleedbox, page.artbox):
        b.lower_left = (float(box.left), top - H)
        b.upper_right = (float(box.left) + W, top)
writer.write(src)
