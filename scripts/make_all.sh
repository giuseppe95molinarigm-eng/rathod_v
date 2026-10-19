#!/usr/bin/env bash
# Build every deliverable: print manuscript (DOCX + PDF), Kindle EPUB,
# printables PDF and covers sized to the final page count.
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
(cd "$ROOT/cover" && node render.js "$PAGES")
cp "$ROOT/cover/paperback_cover_${PAGES}p.pdf" "$OUTDIR/${NAME}_-_Paperback_Cover_${PAGES}p.pdf"
cp "$ROOT/cover/kindle_cover.jpg" "$OUTDIR/${NAME}_-_Kindle_Cover.jpg"

python3 "$ROOT/scripts/build_epub.py" "$OUTDIR/${NAME}_-_Kindle.epub" "$ROOT/cover/kindle_cover.jpg"

WORK="$(mktemp -d)"
cp "$ROOT"/printables/*.md "$WORK/"
cp "$ROOT/manuscript/52_ch18.md" "$ROOT/manuscript/71_appendices.md" "$WORK/"
sed -i 's/^These tools are designed to be used\. .*/These tools are designed to be used: print as many copies as you like./' "$WORK/71_appendices.md"
BOOK_SRC="$WORK" python3 "$ROOT/scripts/build_book.py" "$WORK/${NAME}_-_Printables.docx" >/dev/null
soffice --headless --convert-to "$PDF_FILTER" --outdir "$OUTDIR" "$WORK/${NAME}_-_Printables.docx" >/dev/null 2>&1
rm -rf "$WORK" "$OUTDIR/${NAME}_-_Printables.toc.json"
echo "interior pages (cover sized for): $PAGES"
ls -1 "$OUTDIR"
