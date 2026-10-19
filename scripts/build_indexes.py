"""Generate the Recipe Index and Symptom Index (manuscript/85_*.md, 86_*.md).

Page numbers are filled in by the two-pass build: {{r:Name}} resolves to the
page where a recipe or snack starts, {{p:Text}} to the first page after
Part One that contains Text.
"""

import os
import re

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
MS = os.path.join(ROOT, "manuscript")

CHAPTERS = [("41_ch13.md", "Breakfasts"), ("42_ch14.md", "Lunches"),
            ("43_ch15.md", "Dinners"), ("44_ch16.md", "Snacks and Sips")]

recipes = []   # (course, name, tags)
for fname, course in CHAPTERS:
    cur = None
    for line in open(os.path.join(MS, fname), encoding="utf-8"):
        m = re.match(r"@(recipe|snack) (.+)", line)
        if m:
            cur = m.group(2).strip()
            continue
        if cur and line.startswith(("@tags ", "@info ")):
            tags = [t.strip() for t in line.split(" ", 1)[1].split("|")]
            recipes.append((course, cur, tags))
            cur = None

LABELS = [
    ("Gluten-free (GF)", "GF"), ("Dairy-free (DF)", "DF"), ("Vegetarian (V)", "V"),
    ("Vegan (VG)", "VG"), ("Make-Ahead", "Make-Ahead"),
    ("Ready in 30 minutes or less", None),
]

out = ["@front Recipe Index",
       "Page numbers refer to the first page of each recipe. Labels marked "
       "\"option\" in a recipe need a simple swap, which you'll find in its "
       "Swaps & Variations box; those recipes are listed under the label too.",
       ""]
out.append("## By Meal")
out.append("")
for _, course in CHAPTERS:
    out.append(f"### {course}")
    out.append("")
    for c, name, _ in recipes:
        if c == course:
            out.append(f"- {name}, {{{{r:{name}}}}}")
    out.append("")

out.append("## By Label")
out.append("")
for title, code in LABELS:
    names = []
    for _, name, tags in recipes:
        if code is None:
            if any(re.fullmatch(r"\d+ min", t) for t in tags):
                names.append(name)
        elif any(t == code or t == f"{code} option" for t in tags):
            names.append(name)
    out.append(f"### {title}")
    out.append("")
    for name in sorted(names, key=str.lower):
        out.append(f"- {name}, {{{{r:{name}}}}}")
    out.append("")
open(os.path.join(MS, "85_recipe_index.md"), "w", encoding="utf-8").write("\n".join(out))

SYMPTOMS = [
    ("Anxiety and low mood", ["Why Your Mood Changes With Your Hormones", "Feeding a steadier mind"]),
    ("Bladder leaks and the pelvic floor", ["The Muscles Nobody Talks About"]),
    ("Bloating and digestion", ["Hot, Achy, Bloated", "Calming an unsettled gut"]),
    ("Bone loss and osteoporosis", ["Your Bones Are Alive", "Strength Training: The Habit That Pays You Back",
                                    "Avoid rounding your spine under load"]),
    ("Brain fog and memory", ["Keeping Your Brain Sharp", "Where Did That Word Go?",
                              "Forgetfulness, trouble concentrating"]),
    ("Cravings and emotional eating", ["The 10 a.m. Crash", "Why the 4 p.m. cookie calls your name"]),
    ("Energy and fatigue", ["The 10 a.m. Crash", "Fueling a full day"]),
    ("Headaches and migraines", ["Hot, Achy, Bloated"]),
    ("Heart health", ["Your Heart Needs You Now", "Moving for Your Heart"]),
    ("Heavy or irregular periods", ["The First Signs Are Easy to Miss", "When to Pick Up the Phone"]),
    ("Hot flashes and night sweats", ["Why your internal thermostat goes haywire", "The Symptoms That Can Peak Now",
                                      "What really happened in 2002", "When Hormones Aren't Right for You",
                                      "Changing How Hot Flashes Feel", "What tends to light the fuse"]),
    ("Irritability, rage and tears", ["When Big Feelings Feel Bigger"]),
    ("Joint aches and stiffness", ["Joint pain and stiffness are common", "Joint Aches and Stiffness"]),
    ("Low libido", ["Desire can change too"]),
    ("Palpitations, itchy skin, tinnitus and frozen shoulder", ["The Symptoms Nobody Warned You About"]),
    ("Skin, hair and nails", ["Skin, Hair and Nails"]),
    ("Sleep problems", ["Sleep often changes before anything else", "Sleep: Cooler, Darker, Calmer",
                        "Caffeine curfews, evening meals and sleepy foods"]),
    ("Stress", ["Stress and the Art of Doing Nothing"]),
    ("Vaginal dryness and painful sex", ["Desire, Comfort and Connection", "Comfort Where It Counts"]),
    ("Weight and body shape", ["Your Metabolism Hasn't Broken", "The Metabolism Myth", "Why the Middle Changes Shape"]),
]
out = ["@front Symptom Index",
       "Find it fast: the pages where each symptom is explained, and where you'll "
       "find practical help. For food strategies symptom by symptom, start with "
       "Chapter 10.",
       ""]
for term, targets in SYMPTOMS:
    refs = ", ".join(f"{{{{p:{t}}}}}" for t in targets)
    out.append(f"- **{term}:** {refs}")
out.append("")
open(os.path.join(MS, "86_symptom_index.md"), "w", encoding="utf-8").write("\n".join(out))
print(len(recipes), "recipes and snacks indexed")
