"""Build the reflowable Kindle EPUB from manuscript/*.md.

Usage: python3 build_epub.py OUT.epub [cover.jpg]

Kindle-specific choices: boxes become styled text (not tables), the contents
list is clickable, index entries link to their sections, meal-plan tables
become lists, write-in tables point to the printable PDF, and images carry
alt text.
"""

import glob
import os
import re
import subprocess
import sys
import tempfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(ROOT, "manuscript")
OUT = sys.argv[1]
COVER = sys.argv[2] if len(sys.argv) > 2 else None

TITLE = "The Menopause Diet for Busy Women"
SUBTITLE = ("Quick, Delicious Recipes and a Simple Meal Plan to Balance Hormones, "
            "Boost Energy and Feel Like Yourself Again")
AUTHOR = "Neha Menon"
PRINTABLES = "[website address to be added]"

LABEL = {"summary": "Chapter Summary", "try": "Try This Week",
         "evidence": "What the Evidence Says", "note": "Good to Know",
         "important": "Important", "why": "Why It Helps",
         "swaps": "Swaps & Variations"}


def slug(text, prefix=""):
    s = re.sub(r"[^a-z0-9]+", "-", text.lower()).strip("-")
    return prefix + s


def inline(text):
    text = text.replace("[ ] ", "☐ ")
    text = re.sub(r"==(.+?)==", r"\1", text)
    return text


files = sorted(glob.glob(os.path.join(SRC, "*.md")))
sections = []          # (id, heading text, list of body lines) for index links
heading_ids = {}

# first pass: collect heading ids so index entries can link to them
for path in files:
    cur_id = None
    for line in open(path, encoding="utf-8"):
        line = line.rstrip("\n")
        m = re.match(r"@(recipe|snack) (.+)", line)
        if m:
            heading_ids[m.group(2).strip()] = slug(m.group(2), "r-")
        m = re.match(r"@chapter (\d+) \| (.+)", line)
        if m:
            cur_id = slug(m.group(2), f"ch{m.group(1)}-")
            sections.append([cur_id, m.group(2), []])
            continue
        m = re.match(r"(#{2,3}) (.+)", line)
        if m and cur_id:
            hid = slug(m.group(2), cur_id + "-")
            sections.append([hid, m.group(2), []])
            continue
        if sections:
            sections[-1][2].append(line)


def link_for_text(text):
    """Find the section (id, heading) that first contains a symptom-index target."""
    for hid, head, body in sections:
        if text in head or any(text in b for b in body):
            return hid, head
    return None, None


def convert_table(rows):
    """Return markdown for a table, adapted for Kindle."""
    header, body = rows[0], rows[1:]
    header = [re.sub(r"@[\d.]+$", "", h) for h in header]
    blank = body and all(all(not c.strip() for c in r[1:]) for r in body)
    if blank:
        return [f"*This is a write-in table. Download the printable version at {PRINTABLES}.*", ""]
    if header[0] == "Day" and any("<br>" in c for r in body for c in r):
        out = []
        for r in body:
            day = r[0].replace(" <br> ", ", ")
            out.append(f"**{day}**")
            out.append("")
            for part in r[1].split(" <br> "):
                out.append(f"- {part}")
            out.append("")
        return out
    out = ["| " + " | ".join(header) + " |", "|" + "---|" * len(header)]
    for r in body:
        r = r + [""] * (len(header) - len(r))
        out.append("| " + " | ".join(c.replace(" <br> ", "; ") for c in r) + " |")
    return out + [""]


md = []
USED = {}
md.append(f"# Copyright {{.unnumbered}}")
for path in files:
    lines = open(path, encoding="utf-8").read().splitlines()
    head, rest = lines[0].strip(), lines[1:]
    if head in ("@title", "@contents"):
        continue
    if head == "@copyright":
        for l in rest:
            if l.strip():
                md.append(inline(l))
                md.append("")
        continue
    if head.startswith("@front "):
        name = head[7:].strip()
        md.append(f"# {name} {{#{slug(name, 'fm-')} .unnumbered}}")
    elif head.startswith("@part "):
        num, title, *intro = [x.strip() for x in head[6:].split("|")]
        md.append(f"# Part {num}: {title} {{#{slug(title, 'part-')} .unnumbered .part}}")
        md.append("")
        for t in intro:
            md.append(f"::: tagline\n{t}\n:::")
            md.append("")
    elif head.startswith("@chapter "):
        num, title = [x.strip() for x in head[9:].split("|")]
        md.append(f"# Chapter {num}: {title} {{#{slug(title, 'ch' + num + '-')} .unnumbered}}")
    md.append("")

    i = 0
    chapter_prefix = ""
    m = re.match(r"@chapter (\d+) \| (.+)", head)
    if m:
        chapter_prefix = slug(m.group(2), f"ch{m.group(1)}-")
    while i < len(rest):
        line = rest[i].rstrip()
        s = line.strip()
        if s == "@newpage":
            i += 1
            continue
        if line.startswith(":::"):
            kind_title = line[3:].strip()
            kind, _, title = kind_title.partition(" ")
            label = LABEL.get(kind, "")
            if title and kind in ("evidence", "note", "important") and title.lower() != label.lower():
                caption = f"{label}: {title}"
            else:
                caption = title or label
            md.append(f"::: {{.box .{kind}}}")
            md.append(f"**{caption.upper()}**")
            md.append("")
            i += 1
            while i < len(rest) and not rest[i].startswith(":::"):
                b = rest[i].rstrip()
                if b.strip():
                    md.append(inline(b))
                    md.append("" if not b.startswith(("- ", "[ ] ")) else "")
                i += 1
            md.append(":::")
            md.append("")
            i += 1
            continue
        if line.startswith("|"):
            rows = []
            while i < len(rest) and rest[i].startswith("|"):
                cells = [c.strip() for c in rest[i].strip().strip("|").split("|")]
                if not all(re.fullmatch(r":?-{2,}:?", c) for c in cells):
                    rows.append(cells)
                i += 1
            md.extend(convert_table(rows))
            continue
        if line.startswith("@recipe ") or line.startswith("@snack "):
            name = line.split(" ", 1)[1].strip()
            level = "##" if line.startswith("@recipe") else "###"
            md.append(f"{level} {name} {{#{heading_ids[name]}}}")
        elif line.startswith(("@tags ", "@info ")):
            tags = " · ".join(t.strip() for t in line.split(" ", 1)[1].split("|"))
            md.append(f"::: tags\n{tags}\n:::")
        elif line.startswith("@meta "):
            meta = " · ".join(t.strip() for t in line[6:].split("|"))
            md.append(f"::: meta\n{meta}\n:::")
        elif line.startswith(">> "):
            md.append(f"::: bridge\n{inline(line[3:])}\n:::")
        elif line.startswith("%% "):
            md.append(f"::: scene\n{inline(line[3:])}\n:::")
        elif line.startswith("> "):
            md.append(f"::: tagline\n{inline(line[2:])}\n:::")
        elif line.startswith("#### "):
            md.append(f"**{line[5:]}**")
        elif re.match(r"#{2,3} ", line):
            hashes, text = line.split(" ", 1)
            hid = slug(text, chapter_prefix + "-") if chapter_prefix else slug(text, "s-")
            n = USED.get(hid, 0)
            USED[hid] = n + 1
            if n:
                hid = f"{hid}-{n + 1}"
            md.append(f"{hashes} {text} {{#{hid}}}")
        elif line.startswith("* "):
            md.append("- " + inline(line[2:]))
        elif line.startswith("[ ] "):
            md.append("- " + inline(line))
        elif "{{" in line:
            def rep(m):
                kind, target = m.group(1), m.group(2).strip()
                if kind == "r":
                    return f"[page]({'#' + heading_ids[target]})"
                hid, head = link_for_text(target)
                return f"[{head}](#{hid})" if hid else target
            entry = re.sub(r"\{\{([rp]):([^}]+)\}\}", rep, line)
            # recipe index: link the recipe name itself
            m2 = re.match(r"- (.+), \[page\]\((#[^)]+)\)$", entry)
            if m2:
                entry = f"- [{m2.group(1)}]({m2.group(2)})"
            md.append(inline(entry))
        else:
            md.append(inline(line))
        is_list = line.startswith(("- ", "* ", "[ ] ")) or bool(re.match(r"\d+\. ", line)) or line.startswith("  ")
        if not is_list:
            md.append("")
        elif i + 1 < len(rest):
            nxt = rest[i + 1]
            if not (nxt.startswith(("- ", "* ", "[ ] ", "  ")) or re.match(r"\d+\. ", nxt)):
                md.append("")
        i += 1
    md.append("")

# printable note for Kindle readers near the front
text = "\n".join(md)
text = text.replace(
    "Page numbers refer to the first page of each recipe.",
    "Tap a recipe name to go straight to it.")
text = text.replace(
    "Find it fast: the pages where each symptom is explained",
    "Find it fast: the sections where each symptom is explained")
text = text.replace(
    "Paperback readers can photocopy them or write in the book; Kindle readers can download",
    "Kindle readers can download")

CSS = """
body { font-family: serif; line-height: 1.45; }
h1 { color: #5A3E5D; font-family: sans-serif; margin-top: 1.2em; page-break-before: always; }
h1.part { text-align: center; margin-top: 3em; }
h2 { color: #5A3E5D; font-family: sans-serif; margin-top: 1.2em; }
h3 { font-family: sans-serif; margin-top: 1em; }
.tagline { font-style: italic; color: #555; margin: 0.6em 0 1em 0; }
.scene { font-style: italic; margin: 0.8em 1.2em; }
.bridge { font-style: italic; color: #5A3E5D; text-align: center; margin-top: 1.4em; }
.box { border-left: 4px solid #5A3E5D; background: #F3EFF3; padding: 0.4em 0.8em; margin: 1em 0; }
.box.try { background: #F6F1EA; }
.box.evidence, .box.why { background: #EEF2F0; }
.box.important { background: #F7ECEC; }
.tags { font-family: sans-serif; font-size: 0.8em; letter-spacing: 0.05em; color: #555; text-transform: uppercase; }
.meta { font-family: sans-serif; font-size: 0.85em; background: #F3EFF3; padding: 0.3em 0.6em; margin: 0.4em 0 0.8em 0; }
table { border-collapse: collapse; margin: 0.6em 0; font-size: 0.9em; }
th { background: #5A3E5D; color: #fff; padding: 0.3em; text-align: left; }
td { border-bottom: 1px solid #ccc; padding: 0.3em; vertical-align: top; }
"""

with tempfile.TemporaryDirectory() as tmp:
    mdp = os.path.join(tmp, "book.md")
    open(mdp, "w", encoding="utf-8").write(text)
    css = os.path.join(tmp, "kindle.css")
    open(css, "w").write(CSS)
    meta = os.path.join(tmp, "meta.yaml")
    open(meta, "w").write(
        f"---\ntitle:\n- type: main\n  text: \"{TITLE}\"\n- type: subtitle\n  text: \"{SUBTITLE}\"\n"
        f"creator:\n- role: author\n  text: \"{AUTHOR}\"\nlang: en-US\n"
        f"rights: \"Copyright 2026 {AUTHOR}. All rights reserved.\"\n---\n")
    cmd = ["pandoc", meta, mdp, "-f", "markdown+fenced_divs+bracketed_spans",
           "-t", "epub3", "-o", OUT, "--toc", "--toc-depth=1",
           "--css", css, "--split-level=1", "-M", "toc-title=Contents"]
    if COVER:
        cmd += ["--epub-cover-image", COVER]
    subprocess.run(cmd, check=True)

# give the cover image alt text (the only image in the ebook)
import shutil
import zipfile
ALT = os.environ.get("COVER_ALT", f"Cover of {TITLE} by {AUTHOR}")
tmp_epub = OUT + ".tmp"
with zipfile.ZipFile(OUT) as zin, zipfile.ZipFile(tmp_epub, "w") as zout:
    for item in zin.infolist():
        data = zin.read(item.filename)
        if item.filename.endswith("cover.xhtml"):
            t = data.decode("utf-8")
            t = re.sub(r"<svg ", f'<svg role="img" aria-label="{ALT}" ', t, count=1)
            t = re.sub(r"(<svg[^>]*>)", r"\1<title>" + ALT + "</title>", t, count=1)
            t = re.sub(r'<img ([^>]*?)alt="[^"]*"', r'<img \1', t)
            t = re.sub(r"<img ", f'<img alt="{ALT}" ', t)
            data = t.encode("utf-8")
        compress = zipfile.ZIP_STORED if item.filename == "mimetype" else zipfile.ZIP_DEFLATED
        zout.writestr(item, data, compress_type=compress)
shutil.move(tmp_epub, OUT)
print("saved", OUT)
