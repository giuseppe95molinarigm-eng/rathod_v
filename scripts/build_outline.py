"""Build the proposed Table of Contents / detailed outline for
"The Menopause Diet for Busy Women" as a 6 x 9 in Word document."""

import sys
from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

OUT = sys.argv[1] if len(sys.argv) > 1 else "outline.docx"

TITLE = "The Menopause Diet for Busy Women"
SUBTITLE = ("Easy Recipes, a 12-Week Meal Plan, and Simple Daily Practices "
            "to Support Weight Management, Energy, and Wellbeing")

BODY_FONT = "Cambria"
HEAD_FONT = "Calibri"
ACCENT = RGBColor(0x5A, 0x3E, 0x5D)   # muted plum, prints as dark grey
GREY = RGBColor(0x59, 0x59, 0x59)

# ---------------------------------------------------------------------------
# Outline content (revision 2: client tracked changes and decisions applied)
# Each chapter: (number, title, est_pages, summary, [(section, [points])])
# ---------------------------------------------------------------------------

FRONT = [
    ("Title Page, Copyright & Medical Disclaimer", 3,
     "Full title and subtitle, copyright notice, and a clear statement that "
     "the book is educational and not a substitute for individual medical "
     "advice."),
    ("Contents", 3, "Clickable table of contents (Kindle) / paginated (paperback)."),
    ("A Letter to You", 2,
     "A warm, personal welcome: you are not broken, you are changing. What "
     "this book offers and what it will never do (shame, rigid rules, "
     "miracle promises)."),
    ("How to Use This Book", 3,
     "Three ways in: read cover to cover, jump to your symptom, or start the "
     "12-week plan straight away. Recipe labels key (GF, DF, V, VG, "
     "Make-Ahead, 30 Minutes or Less) and how the sections connect."),
    ("A Note on the Science", 1,
     "How the book uses evidence and how to read the “established” "
     "and “emerging” labels used throughout."),
]

PARTS = [
    ("Part One", "Understanding Your Changing Body", [
        (1, "The Menopause Transition, Explained Simply", 9,
         "A clear, reassuring foundation so every later recommendation makes sense.",
         [("What menopause actually is (and isn’t)",
           ["One day in time vs. a transition of several years",
            "Average age, natural, surgical and medically induced menopause"]),
          ("Early, surgical and medical menopause: when your timeline is different", []),
          ("Meet your hormones",
           ["Estrogen, progesterone, testosterone, FSH and cortisol in plain language",
            "Why fluctuation, not just decline, drives many symptoms"]),
          ("The three stages at a glance",
           ["Perimenopause, menopause, post-menopause: timeline graphic and summary table"]),
          ("Why every woman’s experience is different",
           ["Genetics, lifestyle, stress, culture and health history"]),
          ("Not one-size-fits-all: culture, ethnicity and your menopause", []),
          ("Chapter summary & key takeaways", [])]),
        (2, "Perimenopause: The Shifting Years", 9,
         "The often-confusing years before the final period, when symptoms can "
         "begin long before women expect them.",
         [("Early signs you might be in perimenopause",
           ["Cycle changes, sleep disruption, anxiety, heavier or lighter periods"]),
          ("Common physical changes",
           ["Hot flashes, night sweats, breast tenderness, headaches, bloating"]),
          ("The surprising symptoms no one warned you about: palpitations, "
           "itchy skin, tinnitus, frozen shoulder", []),
          ("Energy, cravings and the blood-sugar connection", []),
          ("What your body needs most in this stage",
           ["Iron (if periods are heavy), B vitamins, magnesium, steady protein"]),
          ("Is it perimenopause, or something else?",
           ["Thyroid changes, low iron and depression: overlapping signs worth checking"]),
          ("When to talk to your doctor", []),
          ("Perimenopause checklist & chapter summary", [])]),
        (3, "Menopause: Reaching the Milestone", 8,
         "The 12-months-without-a-period marker, and what it means for daily life.",
         [("How menopause is confirmed",
           ["Why blood tests are not always needed or reliable"]),
          ("The symptoms that may peak now: hot flashes, sleep and joints", []),
          ("Desire, comfort and connection: your sexual health",
           ["Libido, vaginal dryness and comfortable intimacy"]),
          ("The metabolism shift and changes in body composition", []),
          ("Skin, hair and nails: what’s changing and what helps", []),
          ("What your body needs most in this stage",
           ["Protein, calcium, vitamin D, fiber, phytoestrogen-rich foods"]),
          ("Menopause checklist & chapter summary", [])]),
        (4, "Post-Menopause: Thriving Beyond", 8,
         "Life after the transition is a new chapter, not an ending.",
         [("Which symptoms ease and which may linger", []),
          ("Protecting your bones", ["Osteoporosis risk and what helps"]),
          ("Protecting your heart", ["Cholesterol, blood pressure and lifestyle"]),
          ("Muscle, balance and healthy aging", []),
          ("Brain health and staying sharp", []),
          ("What your body needs most in this stage",
           ["Vitamin B12, calcium, vitamin D, omega-3s, higher protein"]),
          ("Know your numbers: the health checks worth asking for",
           ["Bone density scan, blood pressure, cholesterol and blood sugar"]),
          ("Post-menopause checklist & chapter summary", [])]),
        (5, "Mind and Mood: The Emotional Side of Menopause", 7,
         "Validating the psychological changes that are too often dismissed.",
         [("Anxiety, low mood and irritability: why they happen", []),
          ("When big feelings feel bigger: rage, tears and overwhelm", []),
          ("Brain fog and memory lapses", []),
          ("Identity, body image and self-compassion", []),
          ("Relationships, work and talking about menopause", []),
          ("When to seek support", ["Red flags and where to turn"]),
          ("Chapter summary & reflection prompts", [])]),
    ]),
    ("Part Two", "Treatments, Therapies & Lifestyle", [
        (6, "Treatments and Therapies: Knowing Your Options", 10,
         "Balanced, evidence-informed overview to help readers have informed "
         "conversations with their healthcare provider. Clearly labels "
         "established evidence vs. emerging findings.",
         [("Menopausal hormone therapy (MHT/HRT)",
           ["Types, benefits, risks and who it may suit"]),
          ("Vaginal estrogen and genitourinary health", []),
          ("Non-hormonal prescription options",
           ["When hormone therapy isn’t an option for you"]),
          ("Cognitive behavioral therapy (CBT) for hot flashes and sleep", []),
          ("Supplements: what’s supported, what’s unproven, what to be careful with",
           ["Creatine, collagen and magnesium: hype or help?"]),
          ("Feeling unheard? How to advocate for yourself at appointments", []),
          ("Questions to ask your doctor (printable checklist)", []),
          ("Chapter summary", [])]),
        (7, "Everyday Habits That Make a Difference", 8,
         "The daily habits with the strongest evidence for midlife wellbeing.",
         [("Strength training made simple: protecting muscle and bone", []),
          ("Cardio and everyday movement for your heart", []),
          ("Sleep hygiene for hot nights and busy minds", []),
          ("Stress, cortisol and the power of rest", []),
          ("Rethinking your drinks: alcohol and caffeine",
           ["Plus: support if you want to stop smoking"]),
          ("Pelvic floor health", []),
          ("Connection, community and purpose", []),
          ("Lifestyle checklist & chapter summary", [])]),
    ]),
    ("Part Three", "Nourishing Your Changing Body", [
        (8, "Not Sure of Your Stage? Start Here", 3,
         "A clear way into the nutrition chapters, including for readers who "
         "skipped Part One.",
         [("Side-by-side table: perimenopause vs. menopause vs. post-menopause",
           ["Priority nutrients, focus areas, symptom focus, foods to emphasize"]),
          ("Meal timing and hydration at each stage", []),
          ("Where to go next in the book for your stage", [])]),
        (9, "The Menopause Plate: Your Everyday Framework", 10,
         "A simple, visual way to build balanced meals without counting every calorie.",
         [("If food has felt complicated: a gentle note before we begin", []),
          ("The plate method",
           ["Half vegetables, a quarter protein, a quarter fiber-rich carbs, plus healthy fats"]),
          ("Protein: your midlife power nutrient",
           ["How much, how to spread it across the day, plant and animal sources"]),
          ("Fiber and the gut-hormone connection",
           ["Meet the estrobolome: how your gut helps process estrogen"]),
          ("Healthy fats and omega-3s", []),
          ("Phytoestrogens: soy, flax and legumes, what the evidence says",
           ["Soy and breast cancer: clearing up the confusion"]),
          ("Key micronutrients",
           ["Calcium, vitamin D, magnesium, iron, B vitamins: food-sources table"]),
          ("Anti-inflammatory eating made simple", []),
          ("Portion guidance with the hand method (no scales needed)", []),
          ("Foods to lean into: quick-reference chart", []),
          ("Chapter summary", [])]),
        (10, "Eating for Your Symptoms", 12,
         "Practical, symptom-by-symptom guidance: what may help, what may "
         "trigger, and simple actions to try.",
         [("Hot flashes and night sweats",
           ["Common triggers (alcohol, caffeine, spicy food, large late meals) and cooling foods"]),
          ("Sleep", ["Evening eating, magnesium and tryptophan-rich foods, caffeine timing"]),
          ("Mood and anxiety", ["Blood-sugar balance, omega-3s, the gut-brain axis"]),
          ("Energy and fatigue", ["Iron, B12, protein at breakfast, hydration"]),
          ("Cravings and emotional eating", ["Why they happen and kind strategies"]),
          ("Brain fog", ["Omega-3s, steady blood sugar, hydration and sleep"]),
          ("Bloating and digestion", ["Pacing fiber, fermented foods and common trigger foods"]),
          ("Joint aches and stiffness",
           ["Anti-inflammatory foods, omega-3s, vitamin D and gentle movement"]),
          ("Each symptom includes: “Try This Week” action box and recipe suggestions", []),
          ("Symptom-to-food quick-reference table", [])]),
        (11, "Rethinking Weight in Midlife", 9,
         "A compassionate, realistic look at weight and body composition in midlife.",
         [("Why weight and shape change during menopause",
           ["Hormones, muscle loss, sleep, stress: it’s not a lack of willpower"]),
          ("Why your shape may shift to the middle", []),
          ("Why crash diets backfire, especially now", []),
          ("Fasting, keto and other trends: what the evidence really says", []),
          ("GLP-1 medications and menopause: eating to protect your muscle", []),
          ("Sustainable strategies that work",
           ["Protein, fiber, strength, sleep, gentle consistency"]),
          ("Measuring progress beyond the scale", []),
          ("Chapter summary", [])]),
        (12, "The Busy Woman’s Kitchen", 11,
         "Everything needed to make healthy eating realistic on a full schedule.",
         [("Meal planning in 20 minutes a week", []),
          ("The menopause-friendly grocery guide",
           ["Printable list by store section: produce, protein, pantry, dairy and alternatives, freezer"]),
          ("Pantry, fridge and freezer staples", []),
          ("Easy additions and swaps", ["Simple swap table, e.g., white bread to seeded rye"]),
          ("Batch cooking and prep-ahead routines",
           ["Cook once, eat three times: how leftovers become new meals"]),
          ("Hydration made easy", []),
          ("Eating out, travel and social occasions", []),
          ("One meal, whole family: cooking for everyone at once", []),
          ("Eating well on a budget", []),
          ("Adapting recipes: gluten-free, dairy-free, vegetarian and vegan", []),
          ("Kitchen checklist & chapter summary", [])]),
    ]),
]

RECIPES = {
    "Breakfasts": ("Protein-rich, mostly make-ahead starts to the day that keep "
                   "energy steady through busy mornings.", [
        ("Golden Flax & Berry Overnight Oats", "V, VG option, GF option, Make-Ahead"),
        ("Spinach & Feta Egg Muffins", "GF, V, Make-Ahead"),
        ("Greek Yogurt Power Bowl with Walnuts & Seeds", "GF, V, 10 min"),
        ("Turmeric Tofu & Veggie Scramble", "GF, DF, VG"),
        ("Besan Chilla with Mint Yogurt", "GF, V, DF option, VG option"),
        ("Green Glow Protein Smoothie", "GF, V, DF option, 5 min"),
        ("Smoked Salmon & Avocado Rye Toast", "DF, GF option"),
        ("Ginger Brown Rice Congee with Soft Egg", "GF, DF, V, Make-Ahead"),
    ]),
    "Lunches": ("Satisfying, portable lunches, including soups and bowls that "
                "pack well for work.", [
        ("Mediterranean Lentil & Roasted Vegetable Salad", "GF, DF option, VG option"),
        ("Sesame Edamame Noodle Bowl", "DF, VG, GF option"),
        ("Black Bean, Corn & Avocado Tacos", "GF, DF, VG, 30 min"),
        ("Tuscan White Bean & Kale Soup", "GF, DF, VG, Make-Ahead"),
        ("Tuna & White Bean Niçoise Salad", "GF, DF"),
        ("Quinoa Tabbouleh with Chickpeas", "GF, DF, VG"),
        ("Miso-Glazed Tofu Buddha Bowl", "DF, VG, GF option"),
        ("Lemon Herb Chicken Grain Bowl", "DF, GF option"),
        ("Spiced Red Lentil Dal with Spinach", "GF, DF, VG, Make-Ahead"),
        ("Smashed Chickpea Salad Sandwich", "DF, VG, GF option"),
    ]),
    "Dinners": ("Family-friendly dinners, many ready in 30 minutes or cooked in "
                "a single pan, pot or slow cooker.", [
        ("Sheet-Pan Salmon with Broccoli & Sweet Potato", "GF, DF, 30 min"),
        ("Chicken Pozole Verde", "GF, DF, Make-Ahead"),
        ("Three-Bean Vegetable Chili", "GF, DF, VG, Make-Ahead"),
        ("Baked Cod with Tomatoes, Olives & Capers", "GF, DF, 30 min"),
        ("Ginger-Garlic Tofu & Vegetable Stir-Fry", "DF, VG, GF option, 30 min"),
        ("Lentil Bolognese with Whole-Grain Pasta", "DF, VG, GF option"),
        ("Greek Turkey Meatballs with Tzatziki", "GF option, DF option"),
        ("Coconut Chickpea & Vegetable Curry", "GF, DF, VG"),
        ("Shrimp & Vegetable Fajita Bowls", "GF, DF, 30 min"),
        ("Quinoa & Black Bean Stuffed Peppers", "GF, V, VG option"),
        ("Slow-Cooker Moroccan Chicken with Apricots", "GF, DF, Make-Ahead"),
        ("West African Peanut & Sweet Potato Stew", "GF, DF, VG, Make-Ahead"),
    ]),
}

SNACKS = {
    "Savory Bites": [
        "Crispy Roasted Chickpeas, Three Ways",
        "Chili-Lime Edamame",
        "Everything-Seasoned Boiled Eggs",
        "Hummus & Rainbow Crudités Box",
        "White Bean & Rosemary Dip",
        "Smoked Salmon Cucumber Bites",
        "Seeded Flax Crackers",
        "Baked Kale Chips",
        "Mini Caprese Skewers",
        "Tamari-Roasted Almonds",
        "Cottage Cheese & Cherry Tomato Bowl",
    ],
    "Sweet Treats": [
        "Apple Slices with Almond Butter & Cinnamon",
        "Walnut-Date Energy Bites",
        "Berry Greek Yogurt Bark",
        "Vanilla Chia Pudding Cups",
        "Dark Chocolate & Almond Clusters",
        "Frozen Yogurt-Dipped Blueberries",
        "Banana-Oat Mini Muffins",
        "Baked Cinnamon Pear with Walnuts",
        "Seeded Rye Crispbread with Ricotta & Pear",
        "Spiced Pumpkin Seed Trail Mix",
    ],
    "Soothing Sips": [
        "Tart Cherry Evening Spritz",
        "Golden Turmeric Latte",
        "Cucumber-Mint Cooling Water",
        "Kefir Berry Shake",
    ],
}

WEEKS = [
    ("Foundations", [
        (1, "Protein-First Breakfasts", "Steady energy and fewer mid-morning cravings"),
        (2, "Plants and Fiber", "Gut health and fullness"),
        (3, "Hydration and Hot-Flash Triggers", "Cooling habits and trigger awareness"),
        (4, "Bones: Calcium and Vitamin D", "Long-term bone protection"),
    ]),
    ("Build", [
        (5, "Balancing Blood Sugar", "Cravings, energy dips and mood"),
        (6, "Omega-3s and Mood", "Emotional wellbeing and brain health"),
        (7, "Sleep-Supporting Evenings", "Lighter dinners and calming routines"),
        (8, "Happy Gut, Less Bloating", "Digestion and comfort"),
    ]),
    ("Thrive", [
        (9, "Plant Power and Phytoestrogens", "Soy, flax and legumes"),
        (10, "Batch-Cook Week", "More meals, less effort"),
        (11, "Real Life: Eating Out and Social Occasions", "Flexibility without guilt"),
        (12, "Your Personal Plan", "Building your own sustainable routine"),
    ]),
]

YOGA = [
    ("Week 1", "Days 1–7", "Gentle Foundations",
     "Breath awareness, grounding poses and body scan meditation"),
    ("Week 2", "Days 8–14", "Strength and Balance",
     "Standing poses and strength sessions for bones, muscles and stability; "
     "focus meditation"),
    ("Week 3", "Days 15–21", "Cool and Calm",
     "Cooling breath, gentle twists and stress-relief meditation"),
    ("Week 4", "Days 22–28", "Restore and Sleep",
     "Restorative poses, evening wind-down and yoga nidra-style relaxation"),
    ("Integration", "Days 29–30", "Your Own Practice",
     "Combining favorites into a routine to continue after the 30 days"),
]

LATER_PARTS = [
    ("Part Four", "The Recipes", [
        (13, "Breakfasts (8 recipes)", 12),
        (14, "Lunches (10 recipes)", 15),
        (15, "Dinners (12 recipes)", 18),
        (16, "25 Simple Snacks", 20),
    ]),
    ("Part Five", "The 12-Week Meal Plan", [
        (17, "How the Meal Plan Works", 4),
        (18, "Weeks 1–12", 36),
    ]),
    ("Part Six", "Your 30-Day Yoga, Strength & Calm Plan", [
        (19, "Getting Started Safely", 4),
        (20, "Pose, Strength & Breathing Library", 12),
        (21, "Days 1–30", 18),
    ]),
]

BACK = [("Closing: A Letter for the Road Ahead", 2),
        ("Appendices: trackers, charts & resources", 8),
        ("Glossary, References & Indexes", 8)]

TOTAL_PAGES = (sum(f[1] for f in FRONT)
               + sum(c[2] for _, _, chs in PARTS for c in chs)
               + sum(c[2] for _, _, chs in LATER_PARTS for c in chs)
               + sum(b[1] for b in BACK))

# ---------------------------------------------------------------------------
# Document helpers
# ---------------------------------------------------------------------------

doc = Document()


def set_font(style, name, size, bold=False, italic=False, color=None):
    style.font.name = name
    style.font.size = Pt(size)
    style.font.bold = bold
    style.font.italic = italic
    if color is not None:
        style.font.color.rgb = color
    rpr = style.element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rfonts.set(qn(attr), name)
    for attr in ("w:asciiTheme", "w:hAnsiTheme", "w:cstheme", "w:eastAsiaTheme"):
        if rfonts.get(qn(attr)) is not None:
            del rfonts.attrib[qn(attr)]


def para_fmt(style, before=0, after=6, line=1.15, keep_next=False, align=None):
    pf = style.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.line_spacing = line
    pf.keep_with_next = keep_next
    pf.widow_control = True
    if align is not None:
        pf.alignment = align


styles = doc.styles
set_font(styles["Normal"], BODY_FONT, 10.5)
para_fmt(styles["Normal"], after=6, line=1.2)

set_font(styles["Title"], HEAD_FONT, 26, bold=True, color=ACCENT)
para_fmt(styles["Title"], after=12, line=1.0, align=WD_ALIGN_PARAGRAPH.CENTER)
styles["Title"].element.get_or_add_pPr()
# remove Title's default bottom border
for el in styles["Title"].element.pPr.findall(qn("w:pBdr")):
    styles["Title"].element.pPr.remove(el)

set_font(styles["Subtitle"], BODY_FONT, 12.5, italic=True, color=GREY)
para_fmt(styles["Subtitle"], after=6, line=1.2, align=WD_ALIGN_PARAGRAPH.CENTER)

set_font(styles["Heading 1"], HEAD_FONT, 17, bold=True, color=ACCENT)
para_fmt(styles["Heading 1"], before=0, after=4, line=1.05, keep_next=True)
set_font(styles["Heading 2"], HEAD_FONT, 12, bold=True, color=RGBColor(0x26, 0x26, 0x26))
para_fmt(styles["Heading 2"], before=10, after=3, line=1.1, keep_next=True)
set_font(styles["Heading 3"], HEAD_FONT, 10.5, bold=True, color=ACCENT)
para_fmt(styles["Heading 3"], before=6, after=2, line=1.1, keep_next=True)

for name in ("List Bullet", "List Bullet 2"):
    set_font(styles[name], BODY_FONT, 10)
    para_fmt(styles[name], after=2, line=1.15)

part_style = styles.add_style("Part Label", 1)
part_style.base_style = styles["Normal"]
set_font(part_style, HEAD_FONT, 11, bold=True, color=GREY)
para_fmt(part_style, before=0, after=2, keep_next=True)

summary_style = styles.add_style("Chapter Summary", 1)
summary_style.base_style = styles["Normal"]
set_font(summary_style, BODY_FONT, 10.5, italic=True, color=GREY)
para_fmt(summary_style, after=6, line=1.2)

small_style = styles.add_style("Small Note", 1)
small_style.base_style = styles["Normal"]
set_font(small_style, BODY_FONT, 9, color=GREY)
para_fmt(small_style, after=4, line=1.15)

# Page setup: 6 x 9 in, mirrored margins
sec = doc.sections[0]
sec.page_width = Inches(6)
sec.page_height = Inches(9)
sec.top_margin = Inches(0.75)
sec.bottom_margin = Inches(0.75)
sec.left_margin = Inches(0.8)    # inside (gutter side)
sec.right_margin = Inches(0.6)   # outside
sec.header_distance = Inches(0.4)
sec.footer_distance = Inches(0.4)
settings = doc.settings.element
mirror = OxmlElement("w:mirrorMargins")
settings.insert(0, mirror)
sec.different_first_page_header_footer = True


def add_page_number(paragraph):
    run = paragraph.add_run()
    for tag, text in (("begin", None), (None, "PAGE"), ("end", None)):
        if tag:
            fc = OxmlElement("w:fldChar")
            fc.set(qn("w:fldCharType"), tag)
            run._r.append(fc)
        else:
            it = OxmlElement("w:instrText")
            it.set(qn("xml:space"), "preserve")
            it.text = text
            run._r.append(it)
    run.font.name = HEAD_FONT
    run.font.size = Pt(9)
    run.font.color.rgb = GREY


fp = sec.footer.paragraphs[0]
fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
add_page_number(fp)
hp = sec.header.paragraphs[0]
hp.alignment = WD_ALIGN_PARAGRAPH.CENTER
hr = hp.add_run(TITLE.upper())
hr.font.name = HEAD_FONT
hr.font.size = Pt(7.5)
hr.font.color.rgb = GREY
hr.font.bold = False
# letter spacing
rpr = hr._r.get_or_add_rPr()
sp = OxmlElement("w:spacing")
sp.set(qn("w:val"), "30")
rpr.append(sp)


def bottom_rule(paragraph, color="5A3E5D", size=8):
    ppr = paragraph._p.get_or_add_pPr()
    bdr = OxmlElement("w:pBdr")
    b = OxmlElement("w:bottom")
    b.set(qn("w:val"), "single")
    b.set(qn("w:sz"), str(size))
    b.set(qn("w:space"), "4")
    b.set(qn("w:color"), color)
    bdr.append(b)
    ppr.append(bdr)


def shade_cell(cell, hex_fill):
    tcpr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_fill)
    tcpr.append(shd)


def cell_margins(table, top=60, bottom=60, left=90, right=90):
    tblpr = table._tbl.tblPr
    mar = OxmlElement("w:tblCellMar")
    for side, val in (("top", top), ("left", left), ("bottom", bottom), ("right", right)):
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:w"), str(val))
        el.set(qn("w:type"), "dxa")
        mar.append(el)
    tblpr.append(mar)


def table_borders(table, color="BFB3C0"):
    tblpr = table._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{side}")
        if side in ("left", "right", "insideV"):
            el.set(qn("w:val"), "nil")
        else:
            el.set(qn("w:val"), "single")
            el.set(qn("w:sz"), "4")
            el.set(qn("w:color"), color)
        borders.append(el)
    tblpr.append(borders)


def repeat_header(row):
    trpr = row._tr.get_or_add_trPr()
    el = OxmlElement("w:tblHeader")
    el.set(qn("w:val"), "true")
    trpr.append(el)


def no_split(row):
    trpr = row._tr.get_or_add_trPr()
    trpr.append(OxmlElement("w:cantSplit"))


def cell_text(cell, text, bold=False, size=9.5, color=None, font=None, italic=False,
              align=None):
    p = cell.paragraphs[0]
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1.1
    if align is not None:
        p.alignment = align
    r = p.add_run(text)
    r.bold = bold
    r.italic = italic
    r.font.size = Pt(size)
    if font:
        r.font.name = font
    if color is not None:
        r.font.color.rgb = color
    return p


def make_table(headers, rows, widths, header_fill="5A3E5D", zebra="F3EFF3"):
    t = doc.add_table(rows=1, cols=len(headers))
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    table_borders(t)
    cell_margins(t)
    tblpr = t._tbl.tblPr
    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "fixed")
    tblpr.append(layout)
    tblw = tblpr.find(qn("w:tblW"))
    tblw.set(qn("w:type"), "dxa")
    tblw.set(qn("w:w"), str(int(sum(w.inches for w in widths) * 1440)))
    for gc, w in zip(t._tbl.tblGrid.findall(qn("w:gridCol")), widths):
        gc.set(qn("w:w"), str(int(w.inches * 1440)))
    hdr = t.rows[0]
    repeat_header(hdr)
    for i, h in enumerate(headers):
        c = hdr.cells[i]
        c.width = widths[i]
        shade_cell(c, header_fill)
        cell_text(c, h, bold=True, size=9, color=RGBColor(0xFF, 0xFF, 0xFF), font=HEAD_FONT)
    for ri, row in enumerate(rows):
        cells = t.add_row().cells
        no_split(t.rows[-1])
        for i, val in enumerate(row):
            c = cells[i]
            c.width = widths[i]
            if zebra and ri % 2 == 1:
                shade_cell(c, zebra)
            if isinstance(val, tuple):
                cell_text(c, val[0], bold=val[1], size=9.5)
            else:
                cell_text(c, val, size=9.5)
    return t


_break_pending = False


def page_break():
    global _break_pending
    _break_pending = True


def chapter_heading(label, title):
    global _break_pending
    p = doc.add_paragraph(label.upper(), style="Part Label")
    p.paragraph_format.page_break_before = _break_pending
    _break_pending = False
    h = doc.add_heading(title, level=1)
    bottom_rule(h)
    return h


def bullets(points, level=1):
    style = "List Bullet" if level == 1 else "List Bullet 2"
    for pt in points:
        doc.add_paragraph(pt, style=style)


# ---------------------------------------------------------------------------
# Title page
# ---------------------------------------------------------------------------

for _ in range(6):
    doc.add_paragraph()
p = doc.add_paragraph("TABLE OF CONTENTS & BOOK OUTLINE", style="Part Label")
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
doc.add_paragraph(TITLE, style="Title")
doc.add_paragraph(SUBTITLE, style="Subtitle")
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
bottom_rule(p, size=6)
for _ in range(3):
    doc.add_paragraph()
info = [
    "Format: 6 × 9 in paperback (KDP) + reflowable Kindle eBook",
    f"Target length: approximately {TOTAL_PAGES - 5}–{TOTAL_PAGES + 10} printed pages",
    "30 main recipes · 25 snacks · 12-week meal plan · 30-day yoga, strength & calm plan",
]
for line in info:
    q = doc.add_paragraph(line, style="Small Note")
    q.alignment = WD_ALIGN_PARAGRAPH.CENTER
q = doc.add_paragraph("Revised outline (v2) for final approval", style="Small Note")
q.alignment = WD_ALIGN_PARAGRAPH.CENTER
q.runs[0].italic = True

# ---------------------------------------------------------------------------
# About this outline
# ---------------------------------------------------------------------------

page_break()
chapter_heading("Before you read", "About This Outline")
doc.add_paragraph(
    "This is the revised structure of the book, updated with your tracked "
    "changes and the decisions we agreed together. Once approved, it becomes "
    "the blueprint for the full manuscript.")
doc.add_paragraph(
    "Page estimates are approximate and refer to the final 6 × 9 in "
    "paperback layout. Recipe names are working titles.")
doc.add_heading("What changed in this version", level=2)
bullets([
    "US English used consistently throughout (fiber, savory, behavioral, favorites).",
    "Treatments, Therapies & Lifestyle now comes straight after Part One, so the "
    "book reads: understand → options → practical chapters.",
    "New opening chapter for the nutrition part: “Not Sure of Your Stage? "
    "Start Here”, with the side-by-side stage table (about 3 pages).",
    "Chapter 11 retitled “Rethinking Weight in Midlife”.",
    "Six recipes swapped for dishes from South Asian, East Asian, Latin American "
    "and West African kitchens; the meal plan follows a “cook once, eat "
    "three times” approach.",
    "Two short no-equipment strength sessions a week added to the 30-day plan.",
    "All new sections from your review added, plus a symptom index.",
    "“A Note on the Science” explains how evidence is used and labeled.",
])
doc.add_heading("Tone and approach", level=2)
bullets([
    "Warm, supportive and empowering, never shame-based",
    "Evidence-informed, written in plain, non-technical language",
    "Clear headings, bullet points, checklists and chapter summaries throughout",
    "Medical disclaimers and regular encouragement to consult a healthcare professional",
])
doc.add_heading("Recipe labels used in the book", level=2)
make_table(
    ["Label", "Meaning"],
    [("GF", "Gluten-free (or “GF option” with a simple swap)"),
     ("DF", "Dairy-free (or “DF option”)"),
     ("V", "Vegetarian"),
     ("VG", "Vegan (or “VG option”)"),
     ("Make-Ahead", "Keeps well for batch cooking or freezing"),
     ("30 min", "Ready in 30 minutes or less (the quickest recipes show their "
                "exact time, e.g., “10 min”)")],
    [Inches(1.1), Inches(3.4)])

# ---------------------------------------------------------------------------
# Contents at a glance
# ---------------------------------------------------------------------------

page_break()
chapter_heading("Overview", "Contents at a Glance")
rows = []
for title, pages, _ in FRONT:
    rows.append(("", title, str(pages)))
for plabel, ptitle, chapters in PARTS:
    part_pages = sum(c[2] for c in chapters)
    rows.append(((plabel, True), (ptitle, True), (str(part_pages), True)))
    for num, title, pages, _, _ in chapters:
        rows.append((f"Ch. {num}", title, str(pages)))
for plabel, ptitle, chapters in LATER_PARTS:
    part_pages = sum(c[2] for c in chapters)
    rows.append(((plabel, True), (ptitle, True), (str(part_pages), True)))
    for num, title, pages in chapters:
        rows.append((f"Ch. {num}", title, str(pages)))
rows.append((("Back", True), ("Back Matter", True), (str(sum(b[1] for b in BACK)), True)))
for title, pages in BACK:
    rows.append(("", title, str(pages)))
rows.append((("", True), ("Estimated total", True), (f"~{TOTAL_PAGES}", True)))
make_table(["", "Section", "Pages"], rows,
           [Inches(0.85), Inches(3.15), Inches(0.6)], zebra=None)

# ---------------------------------------------------------------------------
# Detailed outline
# ---------------------------------------------------------------------------

page_break()
chapter_heading("Detailed outline", "Front Matter")
for title, pages, desc in FRONT:
    doc.add_heading(title, level=2)
    doc.add_paragraph(f"{desc}  (~{pages} page{'s' if pages > 1 else ''})")

for plabel, ptitle, chapters in PARTS:
    page_break()
    chapter_heading(plabel, ptitle)
    for num, title, pages, summary, sections in chapters:
        doc.add_heading(f"Chapter {num}  ·  {title}", level=2)
        doc.add_paragraph(f"{summary}  (~{pages} pages)", style="Chapter Summary")
        for sname, pts in sections:
            doc.add_paragraph(sname, style="List Bullet")
            if pts:
                bullets(pts, level=2)

later = {label: chapters for label, _, chapters in LATER_PARTS}


def pages_of(part, num):
    return next(c[2] for c in later[part] if c[0] == num)


# Part four: recipes
page_break()
chapter_heading("Part Four", "The Recipes")
doc.add_paragraph(
    "Every recipe is written for all cooking levels and follows the same "
    "easy-to-scan layout:")
bullets([
    "Short introduction: why this dish supports you during menopause",
    "Serves · Prep time · Cook time · Dietary labels",
    "Ingredients with clear US and metric quantities",
    "Numbered, step-by-step method",
    "Swaps & Variations: gluten-free, dairy-free, vegetarian and vegan options",
    "Make-ahead and storage tips, including how to turn leftovers into a new meal",
])
doc.add_paragraph(
    "Flavors come from Mediterranean, South Asian, East Asian, Latin American "
    "and West African kitchens, using ingredients found in any large supermarket.")
chapter_no = 13
for course, (blurb, items) in RECIPES.items():
    doc.add_heading(f"Chapter {chapter_no}  ·  {course} ({len(items)} recipes)", level=2)
    doc.add_paragraph(f"{blurb}  (~{pages_of('Part Four', chapter_no)} pages)",
                      style="Chapter Summary")
    make_table(["#", "Recipe", "Labels"],
               [(str(i + 1), name, tags) for i, (name, tags) in enumerate(items)],
               [Inches(0.35), Inches(2.75), Inches(1.5)])
    doc.add_paragraph()
    chapter_no += 1

doc.add_heading("Chapter 16  ·  25 Simple Snacks", level=2)
doc.add_paragraph(
    "Quick, satisfying snacks, ready in 5–15 minutes, to steady energy, "
    "curb cravings and support sleep, each with a note on the nutrients it "
    f"provides.  (~{pages_of('Part Four', 16)} pages)", style="Chapter Summary")
n = 1
for group, items in SNACKS.items():
    doc.add_heading(f"{group} ({len(items)})", level=3)
    make_table(["#", "Snack"], [(str(n + i), s) for i, s in enumerate(items)],
               [Inches(0.35), Inches(4.25)])
    n += len(items)

# Part five: meal plan
page_break()
chapter_heading("Part Five", "The 12-Week Meal Plan")
doc.add_heading("Chapter 17  ·  How the Meal Plan Works", level=2)
doc.add_paragraph(
    "A quick orientation so the plan feels flexible and supportive, never "
    f"rigid.  (~{pages_of('Part Five', 17)} pages)", style="Chapter Summary")
bullets([
    "A ready-made menu for breakfast, lunch and dinner across all 12 weeks: "
    "follow it, borrow from it or ignore it",
    "Built around the book’s 30 recipes with a “cook once, eat three "
    "times” approach: one batch becomes a dinner, a lunch and a quick twist "
    "later in the week",
    "Snack suggestions offered as optional extras",
    "How to adapt the plan to your stage, appetite and dietary preferences",
    "Make it yours: swap any meal, skip any day",
    "Missed a few days? Pick up right where you are",
])
doc.add_heading("Chapter 18  ·  Weeks 1–12", level=2)
doc.add_paragraph(
    "Twelve themed weeks in three phases: Foundations, Build and Thrive.  "
    f"(~{pages_of('Part Five', 18)} pages)", style="Chapter Summary")
doc.add_paragraph("Each week includes:")
bullets([
    "Weekly focus: what we’re working on and why",
    "7-day table with breakfast, lunch and dinner",
    "Grocery list for the week",
    "Sunday prep-ahead plan (60–90 minutes)",
    "End-of-week check-in: symptoms, energy, mood and wins",
])
rows = []
for phase, weeks in WEEKS:
    for wk, theme, focus in weeks:
        rows.append((f"Week {wk}", (theme, True), focus))
    rows[-len(weeks)] = (rows[-len(weeks)][0] + f"\n{phase}",) + rows[-len(weeks)][1:]
make_table(["Week", "Theme", "Focus"], rows,
           [Inches(0.9), Inches(2.0), Inches(1.7)])

# Part six: yoga, strength & calm
page_break()
chapter_heading("Part Six", "Your 30-Day Yoga, Strength & Calm Plan")
doc.add_paragraph(
    "Written, easy-to-follow routines (no equipment or videos needed) to "
    "support muscles, bones, balance, heart health, stress and sleep. Each "
    "day combines a 15–25-minute yoga sequence with a 5–10-minute "
    "meditation or breathing practice. Twice a week, a 15-minute bodyweight "
    "strength session takes the place of the yoga sequence.")
doc.add_heading("Chapter 19  ·  Getting Started Safely", level=2)
doc.add_paragraph(
    "Everything needed to begin with confidence, whatever your fitness "
    f"level.  (~{pages_of('Part Six', 19)} pages)", style="Chapter Summary")
bullets([
    "Benefits of yoga and meditation during menopause",
    "What you need, how to set up your space, how to modify poses",
    "Safety notes for joints, blood pressure and bone health",
    "Beyond the mat: why strength matters and how the twice-weekly sessions work",
])
doc.add_heading("Chapter 20  ·  Pose, Strength & Breathing Library", level=2)
doc.add_paragraph(
    "A step-by-step reference for every pose and practice used in the 30-day "
    f"plan.  (~{pages_of('Part Six', 20)} pages)", style="Chapter Summary")
bullets([
    "Around 25 poses described step by step, with benefits and easier/harder options",
    "Around 8 bodyweight strength moves (squats, wall push-ups, glute bridges "
    "and more), with easier and harder options",
    "Breathing techniques, including cooling breath for hot flashes",
    "Meditation scripts: body scan, loving-kindness, gratitude, sleep relaxation",
])
doc.add_heading("Chapter 21  ·  Days 1–30", level=2)
doc.add_paragraph(
    "Four themed weeks plus two integration days, building gently from "
    "foundations to a practice of your own.  "
    f"(~{pages_of('Part Six', 21)} pages)", style="Chapter Summary")
make_table(["Week", "Days", "Theme", "Focus"],
           [(w, d, (t, True), f) for w, d, t, f in YOGA],
           [Inches(0.85), Inches(0.75), Inches(1.25), Inches(1.75)])
doc.add_paragraph()
doc.add_paragraph(
    "Each day includes: the sequence, time needed, intention for the day, "
    "the meditation, and a simple checkbox to track progress.")
doc.add_paragraph(
    "Strength sessions fall on two set days each week (for example, Days 3 "
    "and 6) and build gradually across the four weeks.")

# Back matter
page_break()
chapter_heading("Back matter", "Closing & Appendices")
doc.add_heading("A Letter for the Road Ahead", level=2)
doc.add_paragraph(
    "An encouraging closing message: what readers have built and how to keep "
    "going.  (~2 pages)")
doc.add_heading("Appendices", level=2)
doc.add_paragraph(
    "Printable tools and quick-reference charts to support the plan.  (~8 pages)",
    style="Chapter Summary")
bullets([
    "Appendix A: Symptom & Mood Tracker (printable)",
    "Appendix B: Weekly Meal Planner Template",
    "Appendix C: Master Grocery List",
    "Appendix D: Measurement & Conversion Charts",
    "Appendix E: Questions to Ask Your Healthcare Provider",
    "Appendix F: Trusted Resources & Organizations",
])
doc.add_heading("Glossary, References & Indexes", level=2)
doc.add_paragraph(
    "Plain-language reference material to help readers find what they need "
    "quickly.  (~8 pages)", style="Chapter Summary")
bullets([
    "Glossary of menopause terms in plain language",
    "References to reputable sources (e.g., The Menopause Society, NHS, NICE, "
    "WHO, peer-reviewed research)",
    "Recipe index by meal type and by dietary label",
    "Find it fast: symptom index",
])

doc.save(OUT)
print("saved", OUT, "estimated pages", TOTAL_PAGES)
