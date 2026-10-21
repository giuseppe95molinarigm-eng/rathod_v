"""List the recipes that need a blank page in front to open on a left page.

Usage: find_pads.py BOOK.pdf BOOK.toc.json OUT.json [--check]

Works from a build made without any padding: within a chapter every recipe
starts on a new page, so each blank page added shifts the rest of that
chapter by one. Chapter openers start on a right-hand page again, which
resets the shift. With --check, report recipes not on an even page instead.
"""
import json
import re
import subprocess
import sys

pdf, toc_path, out = sys.argv[1:4]
check = "--check" in sys.argv
spreads = json.load(open(toc_path))["spreads"]
text = subprocess.run(["pdftotext", "-layout", pdf, "-"], capture_output=True,
                      text=True, check=True).stdout
pages = [re.sub(r"\s+", " ", p) for p in text.split("\f")]
start = next(i for i, p in enumerate(pages) if "PART FOUR" in p and "Recipes" in p)

pads, wrong, shift, chapter, pos = [], [], 0, None, start
for name, ch in spreads:
    n = re.sub(r"\s+", " ", name)
    idx = next(i for i in range(pos, len(pages)) if n in pages[i] and "PROTEIN" in pages[i])
    pos = idx
    if ch != chapter:
        chapter, shift = ch, 0
    physical = idx + 1 + shift
    if physical % 2:
        (wrong if check else pads).append(name)
        shift += 0 if check else 1
if check:
    if wrong:
        print("WARNING: recipes not on a left-hand page:", "; ".join(wrong))
    else:
        print("all", len(spreads), "recipes open on a left-hand page")
else:
    data = json.load(open(out)) if out != "-" else {}
    data["_pad"] = pads
    json.dump(data, open(out, "w"), indent=1)
