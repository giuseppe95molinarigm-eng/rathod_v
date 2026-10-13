"""Find the printed page number of each TOC entry in a rendered PDF.

Front matter is numbered in lower-case roman (physical page = number);
arabic numbering starts at 1 on the first Part opener page.
"""
import json
import re
import subprocess
import sys

pdf, toc_path, out = sys.argv[1:4]
entries = json.load(open(toc_path))
text = subprocess.run(["pdftotext", "-layout", pdf, "-"], capture_output=True,
                      text=True, check=True).stdout
pages = [re.sub(r"\s+", " ", p) for p in text.split("\f")]


def roman(n):
    vals = [(10, "x"), (9, "ix"), (5, "v"), (4, "iv"), (1, "i")]
    res = ""
    for v, s in vals:
        while n >= v:
            res += s
            n -= v
    return res


def needle(label):
    label = re.sub(r"^Part \w+: ", "", label)
    label = re.sub(r"^\d+\. ", "", label)
    return re.sub(r"\s+", " ", label)


# the contents page itself lists every title: start searching after it
start = next(i for i, p in enumerate(pages) if "Contents" in p) + 1
result = {}
body_start = None
pos = start
for tid, label in entries:
    n = needle(label)
    # headings may wrap onto two lines: compare with whitespace normalised
    found = next(i for i in range(pos, len(pages)) if n in pages[i])
    if label.startswith("Part ") and body_start is None:
        body_start = found
    result[tid] = (found, label.startswith("Part ") or body_start is not None)
    pos = found
final = {}
for tid, (idx, body) in result.items():
    final[tid] = str(idx - body_start + 1) if body else roman(idx + 1)
json.dump(final, open(out, "w"), indent=1)
