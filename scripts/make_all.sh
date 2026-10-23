#!/usr/bin/env bash
# Build every deliverable: print manuscript (DOCX + PDF), Kindle EPUB,
# printables PDF, and the cover size the designer needs.
set -euo pipefail
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
OUTDIR="$ROOT/deliverables"
NAME="The_Menopause_Diet_for_Busy_Women"
PDF_FILTER='pdf:writer_pdf_Export:{"IsSkipEmptyPages":{"type":"boolean","value":"false"}}'
export HIGHLIGHT=0

python3 "$ROOT/scripts/build_mealplan.py"
python3 "$ROOT/scripts/build_indexes.py"
"$ROOT/scripts/make_book.sh" "$OUTDIR/${NAME}_-_Paperback_Interior.docx"

PAGES=$(pdfinfo "$OUTDIR/${NAME}_-_Paperback_Interior.pdf" | awk '/Pages/{print $2}')
PAGES=$(( (PAGES + 1) / 2 * 2 ))          # KDP spine is based on an even count

# The covers come from the designer. Put the final (unwatermarked) Kindle
# cover at cover/final/kindle_cover.jpg to embed it in the EPUB; without it
# the EPUB has no cover page and KDP uses the cover uploaded separately.
KINDLE_COVER="$ROOT/cover/final/kindle_cover.jpg"
if [ -f "$KINDLE_COVER" ]; then
  COVER_ALT="Cover of The Menopause Diet for Busy Women by Neha Menon: the title in green and coral above a table of healthy food, with a salmon and quinoa bowl, berries, a green smoothie and chickpeas" \
    python3 "$ROOT/scripts/build_epub.py" "$OUTDIR/${NAME}_-_Kindle.epub" "$KINDLE_COVER"
else
  python3 "$ROOT/scripts/build_epub.py" "$OUTDIR/${NAME}_-_Kindle.epub"
fi

WORK="$(mktemp -d)"
cp "$ROOT"/printables/*.md "$WORK/"
cp "$ROOT/manuscript/71_appendices.md" "$WORK/"
PRINT_OUT="$WORK/52_ch18.md" python3 "$ROOT/scripts/build_mealplan.py" >/dev/null
sed -i 's/^These tools are designed to be used\. .*/These tools are designed to be used: print as many copies as you like./' "$WORK/71_appendices.md"
sed -i '/^@qr$/d' "$WORK/71_appendices.md"
# US Letter (the main printable) and an A4 version
for PAPER in letter a4; do
  OUTNAME="${NAME}_-_Printables_$( [ "$PAPER" = letter ] && echo US_Letter || echo A4 )"
  PAPER=$PAPER BOOK_SRC="$WORK" python3 "$ROOT/scripts/build_book.py" "$WORK/$OUTNAME.docx" >/dev/null
  soffice --headless --convert-to "$PDF_FILTER" --outdir "$OUTDIR" "$WORK/$OUTNAME.docx" >/dev/null 2>&1
  rm -f "$WORK/$OUTNAME.toc.json"
done
rm -rf "$WORK" "$OUTDIR/${NAME}_-_Printables.pdf"
python3 -c "import sys; p=int(sys.argv[1]); print(f'cover spec: {p} pages, spine {p*0.002252:.3f} in, full wrap {12.25+p*0.002252:.3f} x 9.25 in')" "$PAGES"
ls -1 "$OUTDIR"
