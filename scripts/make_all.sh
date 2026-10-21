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
  python3 "$ROOT/scripts/build_epub.py" "$OUTDIR/${NAME}_-_Kindle.epub" "$KINDLE_COVER"
else
  python3 "$ROOT/scripts/build_epub.py" "$OUTDIR/${NAME}_-_Kindle.epub"
fi

WORK="$(mktemp -d)"
cp "$ROOT"/printables/*.md "$WORK/"
cp "$ROOT/manuscript/52_ch18.md" "$ROOT/manuscript/71_appendices.md" "$WORK/"
sed -i 's/^These tools are designed to be used\. .*/These tools are designed to be used: print as many copies as you like./' "$WORK/71_appendices.md"
BOOK_SRC="$WORK" python3 "$ROOT/scripts/build_book.py" "$WORK/${NAME}_-_Printables.docx" >/dev/null
soffice --headless --convert-to "$PDF_FILTER" --outdir "$OUTDIR" "$WORK/${NAME}_-_Printables.docx" >/dev/null 2>&1
python3 "$ROOT/scripts/fix_trim.py" "$OUTDIR/${NAME}_-_Printables.pdf"
rm -rf "$WORK" "$OUTDIR/${NAME}_-_Printables.toc.json"
python3 -c "import sys; p=int(sys.argv[1]); print(f'cover spec: {p} pages, spine {p*0.002252:.3f} in, full wrap {12.25+p*0.002252:.3f} x 9.25 in')" "$PAGES"
ls -1 "$OUTDIR"
