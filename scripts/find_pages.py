"""Find the printed page number of each TOC entry in a rendered PDF.

Front matter is numbered in lower-case roman (physical page = number);
arabic numbering starts at 1 on the first Part opener page.
"""
import json
import re
import subprocess
import sys

pdf, toc_path, out = sys.argv[1:4]
data = json.load(open(toc_path))
entries = data["toc"] if isinstance(data, dict) else data
refs = data.get("refs", []) if isinstance(data, dict) else []
text = subprocess.run(["pdftotext", "-layout", pdf, "-"], capture_output=True,
                      text=True, check=True).stdout
raw_pages = text.split("\f")
pages = [re.sub(r"\s+", " ", p) for p in raw_pages]
# a second copy with end-of-line hyphenation removed, for wrapped titles
dehyph = [re.sub(r"\s+", " ", re.sub(r"(\w)-\s*\n\s*(\w)", r"\1\2", p)) for p in raw_pages]


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
# index page references: r = recipe/snack page, p = first mention after Part One
idx_start = next((i for i, p in enumerate(pages) if "Recipe Index" in p and i > body_start + 20), len(pages))
recipes_start = next(i for i, p in enumerate(pages) if "PART FOUR" in p and i > body_start)
missing = []
for key, kind, target in refs:
    t = re.sub(r"\s+", " ", target)
    hit = None
    for i in range(recipes_start if kind == "r" else body_start, idx_start):
        hay = pages[i] + " " + dehyph[i]
        if t in hay and (kind == "p" or "PROTEIN" in pages[i]):
            hit = i
            break
    if hit is None:
        missing.append(target)
        final[key] = "—"
    else:
        final[key] = str(hit - body_start + 1)
if missing:
    print("WARNING: page references not found:", "; ".join(missing), file=sys.stderr)
json.dump(final, open(out, "w"), indent=1)
