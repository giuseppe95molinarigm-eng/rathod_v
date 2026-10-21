#!/usr/bin/env bash
# Build: pass 0 finds recipes that need a blank page to open on the left,
# pass 1 finds TOC and index page numbers, the final pass fills them in.
set -euo pipefail
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
OUT="${1:-$ROOT/deliverables/The_Menopause_Diet_for_Busy_Women.docx}"
WORK="$(mktemp -d)"
PDF_FILTER='pdf:writer_pdf_Export:{"IsSkipEmptyPages":{"type":"boolean","value":"false"}}'

python3 "$ROOT/scripts/build_book.py" "$WORK/pass0.docx"
soffice --headless --convert-to "$PDF_FILTER" --outdir "$WORK" "$WORK/pass0.docx" >/dev/null 2>&1
echo '{}' > "$WORK/pad.json"
python3 "$ROOT/scripts/find_pads.py" "$WORK/pass0.pdf" "$WORK/pass0.toc.json" "$WORK/pad.json"
python3 "$ROOT/scripts/build_book.py" "$WORK/pass1.docx" "$WORK/pad.json"
soffice --headless --convert-to "$PDF_FILTER" --outdir "$WORK" "$WORK/pass1.docx" >/dev/null 2>&1
python3 "$ROOT/scripts/find_pages.py" "$WORK/pass1.pdf" "$WORK/pass1.toc.json" "$WORK/pages.json"
python3 -c 'import json,sys;d=json.load(open(sys.argv[1]));d.update(json.load(open(sys.argv[2])));json.dump(d,open(sys.argv[1],"w"))' "$WORK/pages.json" "$WORK/pad.json"
python3 "$ROOT/scripts/build_book.py" "$OUT" "$WORK/pages.json"
soffice --headless --convert-to "$PDF_FILTER" --outdir "$(dirname "$OUT")" "$OUT" >/dev/null 2>&1
PDF="${OUT%.docx}.pdf"
python3 "$ROOT/scripts/fix_trim.py" "$PDF"
python3 "$ROOT/scripts/find_pages.py" "$PDF" "${OUT%.docx}.toc.json" "$WORK/check.json"
python3 "$ROOT/scripts/find_pads.py" "$PDF" "${OUT%.docx}.toc.json" - --check
python3 -c 'import json,sys;d=json.load(open(sys.argv[1]));d.update(json.load(open(sys.argv[2])));json.dump(d,open(sys.argv[1],"w"))' "$WORK/check.json" "$WORK/pad.json"
cmp -s <(python3 -c 'import json,sys;print(json.load(open(sys.argv[1])))' "$WORK/pages.json") \
       <(python3 -c 'import json,sys;print(json.load(open(sys.argv[1])))' "$WORK/check.json") \
  && echo "TOC page numbers verified" || echo "WARNING: TOC page numbers changed between passes"
rm -f "${OUT%.docx}.toc.json"
pdfinfo "$PDF" | grep -E "Pages|Page size"
