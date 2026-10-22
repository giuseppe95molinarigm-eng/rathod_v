"""Generate Chapter 18 (the 12-week meal plan) from the recipe chapters.

Daily protein totals and weekly grocery lists are derived from the recipe
files themselves, so the plan stays consistent with the recipes.
Usage: python3 build_mealplan.py  (writes manuscript/52_ch18.md)
"""

import os
import re

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
MS = os.path.join(ROOT, "manuscript")

CODES = {
    # breakfasts
    "OATS": "Golden Flax & Berry Overnight Oats",
    "MUFF": "Spinach & Feta Egg Muffins",
    "BOWL": "Greek Yogurt Power Bowl with Walnuts & Seeds",
    "SCRAM": "Turmeric Tofu & Veggie Scramble",
    "CHILLA": "Besan Chilla with Mint Yogurt",
    "SMOOTH": "Green Glow Protein Smoothie",
    "TOAST": "Smoked Salmon & Avocado Rye Toast",
    "CONGEE": "Ginger Brown Rice Congee with Soft Egg",
    # lunches
    "LENT": "Mediterranean Lentil & Roasted Vegetable Salad",
    "SOBA": "Sesame Edamame Noodle Bowl",
    "TACO": "Black Bean, Corn & Avocado Tacos",
    "SOUP": "Tuscan White Bean & Kale Soup",
    "NIC": "Tuna & White Bean Niçoise Salad",
    "TABB": "Quinoa Tabbouleh with Chickpeas",
    "MISO": "Miso-Glazed Tofu Buddha Bowl",
    "CHICK": "Lemon Herb Chicken Grain Bowl",
    "DAL": "Spiced Red Lentil Dal with Spinach",
    "SAND": "Smashed Chickpea Salad Sandwich",
    # dinners
    "SALMON": "Sheet-Pan Salmon with Broccoli & Sweet Potato",
    "POZ": "Chicken Pozole Verde",
    "CHILI": "Three-Bean Vegetable Chili",
    "COD": "Baked Cod with Tomatoes, Olives & Capers",
    "STIR": "Ginger-Garlic Tofu & Vegetable Stir-Fry",
    "BOLO": "Lentil Bolognese with Whole-Grain Pasta",
    "MEAT": "Greek Turkey Meatballs with Tzatziki",
    "CURRY": "Coconut Chickpea & Vegetable Curry",
    "SHRIMP": "Shrimp & Vegetable Fajita Bowls",
    "PEPPER": "Quinoa & Black Bean Stuffed Peppers",
    "MOROC": "Slow-Cooker Moroccan Chicken with Apricots",
    "PEANUT": "West African Peanut & Sweet Potato Stew",
    # snacks
    "EGGS": "Everything-Seasoned Boiled Eggs",
    "BARK": "Berry Greek Yogurt Bark",
    "CHICKPEAS": "Crispy Roasted Chickpeas, Three Ways",
    "HUMMUS": "Hummus & Rainbow Crudités Box",
    "APPLE": "Apple Slices with Almond Butter & Cinnamon",
    "BITES": "Walnut-Date Energy Bites",
    "COTTAGE": "Cottage Cheese & Cherry Tomato Bowl",
    "BLUEB": "Frozen Yogurt-Dipped Blueberries",
    "EDAMAME": "Chili-Lime Edamame",
    "CAPRESE": "Mini Caprese Skewers",
    "CHIA": "Vanilla Chia Pudding Cups",
    "TRAIL": "Spiced Pumpkin Seed Trail Mix",
    "SALMONBITES": "Smoked Salmon Cucumber Bites",
    "PEAR": "Baked Cinnamon Pear with Walnuts",
    "LATTE": "Golden Turmeric Latte",
    "KEFIR": "Kefir Berry Shake",
    "ALMONDS": "Tamari-Roasted Almonds",
    "RICOTTA": "Seeded Rye Crispbread with Ricotta & Pear",
    "BEANDIP": "White Bean & Rosemary Dip",
}

FLEX_PROTEIN = 25   # a meal out, built around a palm-and-a-half of protein

# Each week: (number, theme, focus paragraph, snacks, days, prep steps)
# A day is (breakfast, lunch, dinner); "LO:X" means leftovers of X.
WEEKS = [
    (1, "Protein-First Breakfasts",
     "This week is about one change: anchoring breakfast with protein. Two make-ahead breakfasts mean there's nothing to decide at 7 a.m., and a big pot of soup covers several lunches. Notice how your energy feels at 10:30 compared with last week.",
     ["EGGS", "BARK", "CHICKPEAS"],
     [("MUFF", "SOUP", "SALMON"), ("OATS", "LO:SALMON", "CHILI"), ("MUFF", "LO:CHILI", "COD"),
      ("OATS", "SOUP", "MEAT"), ("MUFF", "LO:MEAT", "STIR"), ("BOWL", "NIC", "LO:CHILI"),
      ("SCRAM", "SOUP", "PEPPER")],
     ["Bake the Spinach & Feta Egg Muffins and make two jars of overnight oats.",
      "Make the Tuscan White Bean & Kale Soup and portion it for three lunches.",
      "Boil a batch of eggs and roast a tray of chickpeas for snacks.",
      "Freeze the yogurt bark for the week."]),
    (2, "Plants and Fiber",
     "This week we add plants: aim for 30 different ones by Sunday, counting every vegetable, fruit, bean, grain, nut, seed, herb and spice. If you're not used to much fiber, drink plenty of water and let your gut adjust gradually.",
     ["HUMMUS", "APPLE", "BITES"],
     [("OATS", "LENT", "CURRY"), ("SMOOTH", "LO:CURRY", "BOLO"), ("OATS", "LENT", "PEANUT"),
      ("SMOOTH", "LO:BOLO", "SHRIMP"), ("OATS", "LO:PEANUT", "PEPPER"), ("CHILLA", "TABB", "STIR"),
      ("CHILLA", "TABB", "MOROC")],
     ["Roast the vegetables and make the Mediterranean Lentil Salad for three lunches.",
      "Make two jars of overnight oats and portion smoothie packs into freezer bags.",
      "Roll a batch of Walnut-Date Energy Bites.",
      "Make the tabbouleh on Saturday morning for the weekend."]),
    (3, "Hydration and Hot-Flash Comfort",
     "This week focuses on cooling, hydrating meals and on noticing your personal triggers. Keep a pitcher of Cucumber-Mint Cooling Water in the fridge, keep spice on the side, and use your symptom diary to spot patterns. Remember, these are comfort measures; if hot flashes are disrupting your life, see Chapter 6.",
     ["COTTAGE", "BLUEB", "EDAMAME"],
     [("MUFF", "SOBA", "COD"), ("SMOOTH", "TABB", "STIR"), ("MUFF", "LO:STIR", "SALMON"),
      ("BOWL", "TABB", "MEAT"), ("MUFF", "LO:MEAT", "SHRIMP"), ("SMOOTH", "NIC", "CURRY"),
      ("SCRAM", "LO:CURRY", "POZ")],
     ["Bake the egg muffins and make the tabbouleh.",
      "Make the sesame dressing and cook the soba for Monday's lunch.",
      "Freeze a tray of yogurt-dipped blueberries.",
      "Fill a pitcher of Cucumber-Mint Cooling Water."]),
    (4, "Bones: Calcium and Vitamin D",
     "This week every day includes at least two calcium-rich foods, from Greek yogurt and calcium-set tofu to cheese and fortified soy milk, and oily fish brings vitamin D. Pair it with your two strength sessions from Part Six: food and loading work together for your bones.",
     ["CAPRESE", "COTTAGE", "CHIA"],
     [("BOWL", "LO:POZ", "SALMON"), ("MUFF", "MISO", "PEPPER"), ("BOWL", "LO:PEPPER", "STIR"),
      ("MUFF", "LO:STIR", "MEAT"), ("BOWL", "LO:MEAT", "CURRY"), ("TOAST", "SAND", "COD"),
      ("SCRAM", "LO:CURRY", "CHILI")],
     ["Bake the egg muffins and press and cube two blocks of calcium-set tofu.",
      "Make two jars of chia pudding with soy milk.",
      "Cook the quinoa for Tuesday's stuffed peppers and keep it in the fridge.",
      "Check your vitamin D: is a supplement right for you? Ask at your next appointment."]),
    (5, "Balancing Blood Sugar",
     "This week is about steady energy: try not to go much longer than four or five hours without eating, protein and fiber at every meal, and a planned snack for the time of day your cravings usually strike. Try a 10-minute walk after dinner on at least three evenings.",
     ["EGGS", "APPLE", "TRAIL"],
     [("OATS", "LO:CHILI", "MOROC"), ("MUFF", "CHICK", "BOLO"), ("OATS", "CHICK", "SHRIMP"),
      ("MUFF", "CHICK", "LO:BOLO"), ("OATS", "LO:MOROC", "PEANUT"), ("CHILLA", "LO:PEANUT", "SALMON"),
      ("SCRAM", "LENT", "MEAT")],
     ["Make the Lemon Herb Chicken Grain Bowls for three lunches.",
      "Bake the egg muffins and make two jars of overnight oats.",
      "Set up the slow cooker for Monday's Moroccan chicken.",
      "Roast a batch of trail mix and boil eggs."]),
    (6, "Omega-3s and Mood",
     "This week brings oily fish to the table at least twice, plus walnuts, flaxseed and chia every day. Take a moment each evening for the gratitude note from Chapter 5. Omega-3s support heart and brain health over the long term, so notice how you feel without expecting quick changes.",
     ["SALMONBITES", "BITES", "TRAIL"],
     [("OATS", "LO:MEAT", "SALMON"), ("BOWL", "LO:SALMON", "BOLO"), ("OATS", "LO:BOLO", "COD"),
      ("BOWL", "NIC", "CURRY"), ("OATS", "LO:CURRY", "SHRIMP"), ("TOAST", "SOBA", "PEANUT"),
      ("TOAST", "LO:PEANUT", "POZ")],
     ["Make two jars of overnight oats with extra ground flaxseed.",
      "Roll the energy bites and roast the trail mix.",
      "Make the Bolognese sauce and freeze half for a later week.",
      "Hard-boil eggs for Thursday's Niçoise."]),
    (7, "Sleep-Supporting Evenings",
     "This week's dinners are a little lighter and earlier, with a calming bedtime snack if you need one. Try a 2 p.m. caffeine cut-off, swap one evening drink for a Tart Cherry Evening Spritz, and try one of the evening practices from Part Six.",
     ["PEAR", "CHIA", "LATTE"],
     [("MUFF", "LO:POZ", "COD"), ("SMOOTH", "SOUP", "MEAT"), ("MUFF", "LO:MEAT", "STIR"),
      ("SMOOTH", "SOUP", "CURRY"), ("MUFF", "LO:CURRY", "SALMON"), ("CONGEE", "SOUP", "PEPPER"),
      ("CONGEE", "LO:PEPPER", "DAL")],
     ["Make the Tuscan soup for three lunches.",
      "Bake the egg muffins and portion smoothie packs.",
      "Make chia pudding for two evenings.",
      "Start the congee on Saturday morning; it keeps for Sunday too."]),
    (8, "Happy Gut, Less Bloating",
     "This week includes a fermented food on most days (yogurt, kefir, miso) and keeps fiber steady rather than sudden. Eat slowly, take a short walk after meals, and notice which foods leave you comfortable and which don't.",
     ["KEFIR", "CHIA", "EDAMAME"],
     [("CONGEE", "LO:DAL", "SALMON"), ("BOWL", "MISO", "MOROC"), ("CONGEE", "MISO", "COD"),
      ("BOWL", "LO:MOROC", "STIR"), ("OATS", "LO:STIR", "BOLO"), ("OATS", "LO:MOROC", "SHRIMP"),
      ("SCRAM", "LO:BOLO", "CHILI")],
     ["Make a pot of congee for three breakfasts.",
      "Marinate and bake the miso tofu for two lunches.",
      "Set up the slow cooker for Tuesday's Moroccan chicken.",
      "Make chia pudding and keep kefir in the fridge for shakes."]),
    (9, "Plant Power and Phytoestrogens",
     "This week leans on soy and other plant proteins: tofu, edamame, soy milk, chickpeas and lentils, plus a daily spoonful of ground flaxseed. Enjoy them for their protein, fiber and calcium; any effect on hot flashes is a bonus.",
     ["EDAMAME", "CHIA", "HUMMUS"],
     [("OATS", "LO:CHILI", "STIR"), ("SMOOTH", "LO:STIR", "CURRY"), ("OATS", "SOBA", "PEANUT"),
      ("SMOOTH", "LO:PEANUT", "PEPPER"), ("OATS", "SAND", "LO:CURRY"), ("SCRAM", "MISO", "BOLO"),
      ("CHILLA", "LO:BOLO", "CURRY")],
     ["Make two jars of overnight oats with soy milk.",
      "Press two blocks of tofu for the stir-fry and miso bowl.",
      "Make the sesame dressing for Wednesday's noodles.",
      "Portion smoothie packs and make chia pudding."]),
    (10, "Batch-Cook Week",
     "This week does the heavy lifting on Sunday so the weekdays almost cook themselves. Three big batches (pozole, chili and dal) plus make-ahead breakfasts cover most of the week. Freeze any extra portions for the weeks ahead.",
     ["EGGS", "CHICKPEAS", "BARK"],
     [("MUFF", "DAL", "POZ"), ("OATS", "LO:POZ", "CHILI"), ("MUFF", "DAL", "LO:POZ"),
      ("OATS", "LO:CHILI", "SALMON"), ("MUFF", "DAL", "LO:CHILI"), ("BOWL", "LO:SALMON", "MEAT"),
      ("SCRAM", "LO:MEAT", "MOROC")],
     ["Make the Chicken Pozole Verde and the Three-Bean Vegetable Chili.",
      "Make the Spiced Red Lentil Dal for three lunches.",
      "Bake the egg muffins and make two jars of overnight oats.",
      "Boil eggs, roast chickpeas and freeze the yogurt bark."]),
    (11, "Real Life: Eating Out and Social Occasions",
     "This week includes three meals out, because real life does. When you eat out, look for a palm-and-a-half of protein and plenty of plants, enjoy what you choose, and pick up your usual routine at the next meal. One meal never undoes a week.",
     ["ALMONDS", "EGGS", "APPLE"],
     [("OATS", "LO:MOROC", "STIR"), ("MUFF", "LO:STIR", "COD"), ("OATS", "LO:MOROC", "FLEX"),
      ("MUFF", "NIC", "SHRIMP"), ("OATS", "TACO", "FLEX"), ("TOAST", "FLEX", "PEANUT"),
      ("SCRAM", "LO:PEANUT", "SALMON")],
     ["Make two jars of overnight oats and bake the egg muffins.",
      "Roast the almonds and boil eggs: they travel well.",
      "Look up the menus for your meals out and decide in advance if that helps you.",
      "Keep Saturday's peanut stew simple: it makes Sunday's lunch too."]),
    (12, "Your Personal Plan",
     "This final week is built from favorites many women keep on repeat. Use it as a model for your own plan: keep the breakfasts you loved, the batch-cook habits that saved you time, and the snacks that steadied your afternoons. Appendix B has a blank weekly planner.",
     ["RICOTTA", "COTTAGE", "EGGS"],
     [("BOWL", "LENT", "PEPPER"), ("CHILLA", "LO:PEPPER", "CURRY"), ("BOWL", "LO:CURRY", "SALMON"),
      ("OATS", "LENT", "BOLO"), ("OATS", "LO:BOLO", "SHRIMP"), ("CONGEE", "TACO", "POZ"),
      ("SCRAM", "LO:POZ", "MEAT")],
     ["Make the lentil salad for two lunches and stuff the peppers for Monday.",
      "Make two jars of overnight oats.",
      "Boil eggs and buy ricotta and crispbread for snacks.",
      "Sit down with Appendix B and sketch next week's plan."]),
]

PHASES = {1: "Phase 1: Foundations", 5: "Phase 2: Build", 9: "Phase 3: Thrive"}
DAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]

# ---------------------------------------------------------------------------
# Read protein and ingredients from the recipe chapters
# ---------------------------------------------------------------------------

PROTEIN, INGREDIENTS = {}, {}
for name in sorted(os.listdir(MS)):
    if not re.match(r"4[1-6]_ch1[3-6]\.md", name):
        continue
    cur, in_ing = None, False
    for line in open(os.path.join(MS, name), encoding="utf-8"):
        line = line.rstrip("\n")
        m = re.match(r"@(recipe|snack) (.+)", line)
        if m:
            cur, in_ing = m.group(2).strip(), True
            INGREDIENTS[cur] = []
            continue
        if cur and line.startswith(("@meta", "@info")):
            p = re.search(r"Protein ≈?(\d+) g", line)
            PROTEIN[cur] = int(p.group(1)) if p else 0
        if cur and line.startswith("### Method"):
            in_ing = False
        if cur and in_ing and line.startswith("* "):
            INGREDIENTS[cur].append(line[2:])
        if cur and line.startswith("1. "):
            in_ing = False

# fridge life ("Keeps N days" in @meta, None for "Best fresh") and whether
# the recipe says it can be frozen
KEEPS, FREEZES, SNACKS = {}, set(), set()
for name in sorted(os.listdir(MS)):
    if not re.match(r"4[1-6]_ch1[3-6]\.md", name):
        continue
    cur = None
    for line in open(os.path.join(MS, name), encoding="utf-8"):
        m = re.match(r"@(recipe|snack) (.+)", line)
        if m:
            cur = m.group(2).strip()
        elif cur and line.startswith("@meta"):
            k = re.search(r"Keeps (\d+) days?", line)
            KEEPS[cur] = int(k.group(1)) if k else None
        elif cur and line.startswith("@info"):
            KEEPS[cur] = None
            SNACKS.add(cur)
        elif cur and KEEPS.get(cur, 0) is None and cur in SNACKS and \
                re.match(r"\d\. .*([Kk]eep|[Ss]tor)", line):
            # snacks give their fridge life in the method
            k = re.search(r"(?:up to|for) (\d+|a|1) (days?|weeks?)", line)
            if k:
                n = 1 if k.group(1) in ("a", "1") else int(k.group(1))
                KEEPS[cur] = n * (7 if k.group(2).startswith("week") else 1)
        elif cur and re.match(r"- \*\*(Freezing|Storage):\*\*.*\bfreez", line):
            FREEZES.add(cur)

missing = [c for c, n in CODES.items() if n not in PROTEIN]
assert not missing, f"recipes not found: {missing}"

STAPLES = re.compile(
    r"\b(salt|pepper\b(?! flakes)|black pepper|olive oil|neutral oil|oil\b|water|"
    r"cumin|coriander|turmeric|paprika|cinnamon|oregano|chili powder|chili flakes|"
    r"garam masala|curry powder|cayenne|ground ginger|bay leaf|vanilla|baking powder|"
    r"dried thyme|dried rosemary|garlic powder|cocoa|everything-bagel|za'atar|nutritional yeast|ice cubes|\bice\b)",
    re.I)
SECTIONS = [
    ("Protein", r"salmon|cod|shrimp|chicken|turkey|tuna|eggs?\b|tofu|tempeh|halibut"),
    ("Dairy and alternatives", r"yogurt|milk|feta|cheese|ricotta|cottage|mozzarella|kefir|parmesan|cheddar"),
    ("Freezer", r"frozen"),
    ("Produce", r"onion|garlic|ginger|pepper|spinach|kale|broccoli|carrot|celery|tomato|cucumber|"
                r"zucchini|lettuce|arugula|cabbage|potato|lemon|lime|avocado|apple|pear|banana|"
                r"berr|mango|cilantro|parsley|mint|dill|basil|chives|scallion|mushroom|bok choy|"
                r"cauliflower|radish|tomatillo|poblano|jalape|green beans|snow peas|herbs|corn kernels|"
                r"corn \(|squash|kiwi|fruit|vegetables"),
]


def clean(item):
    s = re.sub(r"\([^)]*\)", "", item)                 # drop (metric) notes
    s = re.sub(r"\*\*[^*]+\*\*:?", "", s)               # drop bold variant labels
    s = s.split(",")[0]
    s = re.sub(r"^(about |a |an )", "", s.strip(), flags=re.I)
    s = re.sub(r"^[\d/ .½¼¾-]+", "", s)
    s = re.sub(r"^(cups?|tbsp|tsp|oz|lb|cans?|blocks?|heads?|bunch(es)?|cloves?|stalks?|large|medium|"
               r"small|handfuls?|thumb-sized piece|pinch of|squeeze of|slices?|sprigs?|of)\b\s*", "", s, flags=re.I)
    s = re.sub(r"^(cups?|tbsp|tsp|oz|lb|cans?|blocks?|heads?|bunch(es)?|cloves?|stalks?|large|medium|"
               r"small|handfuls?|of)\b\s*", "", s, flags=re.I)
    s = re.sub(r"^(large|medium|small|of)\s+", "", s, flags=re.I)
    s = re.sub(r"\s+", " ", s).strip(" .;:")
    s = re.sub(r"\b(to serve|optional|or more|plus extra.*)$", "", s, flags=re.I).strip(" ,")
    return s[:1].lower() + s[1:] if s else s


CANON = [
    # (pattern, canonical name, section)
    (r"lemon", "lemons", "Produce"), (r"\blime", "limes", "Produce"),
    (r"\btuna\b", "canned tuna", "Protein"),
    (r"coconut milk", "light coconut milk", "Pantry"),
    (r"crispbread", "seeded rye crispbreads", "Pantry"),
    (r"garlic", "garlic", "Produce"), (r"ginger", "fresh ginger", "Produce"),
    (r"carrot", "carrots", "Produce"), (r"broccoli", "broccoli", "Produce"),
    (r"scallion", "scallions", "Produce"), (r"red onion", "red onions", "Produce"),
    (r"\bonion", "onions", "Produce"), (r"cherry tomato", "cherry tomatoes", "Produce"),
    (r"tomato paste", "tomato paste", "Pantry"),
    (r"(chopped|crushed|diced) tomatoes", "canned tomatoes", "Pantry"),
    (r"bell pepper", "bell peppers", "Produce"), (r"spinach", "baby spinach", "Produce"),
    (r"kale", "kale", "Produce"), (r"cucumber", "cucumbers", "Produce"),
    (r"avocado", "avocados", "Produce"), (r"celery", "celery", "Produce"),
    (r"mushroom", "mushrooms", "Produce"), (r"sweet potato", "sweet potatoes", "Produce"),
    (r"berr", "berries (fresh or frozen)", "Produce"),
    (r"(dill|parsley|chives|basil|cilantro|mint|fresh herbs)", "fresh herbs (dill, parsley, cilantro, mint or basil)", "Produce"),
    (r"lettuce|romaine|little gem|arugula|mixed leaves", "salad leaves", "Produce"),
    (r"\beggs?\b", "eggs", "Protein"), (r"tofu", "tofu (firm and extra-firm, calcium-set if possible)", "Protein"),
    (r"chicken thigh", "boneless, skinless chicken thighs", "Protein"),
    (r"chicken breast|chicken breasts or thighs", "chicken breasts or thighs", "Protein"),
    (r"smoked salmon", "smoked salmon", "Protein"), (r"salmon", "salmon fillets", "Protein"),
    (r"(soy milk|\bmilk\b)", "milk or unsweetened soy milk", "Dairy and alternatives"),
    (r"greek yogurt|plain yogurt|soy yogurt|yogurt", "plain Greek yogurt (or soy yogurt)", "Dairy and alternatives"),
    (r"feta", "feta cheese", "Dairy and alternatives"),
    (r"cannellini|white beans", "cannellini (white) beans, canned", "Pantry"),
    (r"black beans", "black beans, canned", "Pantry"), (r"chickpeas", "chickpeas, canned", "Pantry"),
    (r"mixed beans", "mixed beans, canned", "Pantry"),
    (r"lentils", "lentils", "Pantry"), (r"walnut", "walnuts", "Pantry"),
    (r"almond butter", "almond butter", "Pantry"), (r"almond", "almonds", "Pantry"),
    (r"rolled oats|oats", "rolled oats", "Pantry"), (r"brown rice", "brown rice", "Pantry"),
    (r"quinoa", "quinoa", "Pantry"), (r"maple|honey", "maple syrup or honey", "Pantry"),
    (r"soy sauce|tamari", "low-sodium soy sauce or tamari", "Pantry"),
    (r"stock", "low-sodium stock", "Pantry"), (r"olives", "olives", "Pantry"),
    (r"pita", "whole-grain pitas", "Pantry"), (r"tortilla", "tortillas", "Pantry"),
    (r"dark chocolate", "dark chocolate (70% or higher)", "Pantry"),
    (r"pistachio", "pistachios", "Pantry"), (r"pumpkin seeds", "pumpkin seeds", "Pantry"),
    (r"sesame seeds", "sesame seeds", "Pantry"), (r"edamame", "frozen shelled edamame", "Freezer"),
    (r"corn", "corn (frozen or canned)", "Freezer"), (r"granola", "low-sugar granola", "Pantry"),
    (r"toast|bread|crispbread", "whole-grain or seeded rye bread", "Pantry"),
    (r"pasta|spaghetti|penne", "whole-grain pasta", "Pantry"),
    (r"noodles|soba", "soba or whole-wheat noodles", "Pantry"),
]


def canonical(name):
    low = name.lower()
    for pat, canon, sec in CANON:
        if re.search(pat, low):
            return canon, sec
    return None, None


def grocery(week_codes):
    grouped = {s: set() for s, _ in SECTIONS}
    grouped["Pantry"] = set()
    for code in week_codes:
        for raw in INGREDIENTS.get(CODES[code], []):
            name = clean(raw)
            if name.lower() in ("boneless", "skinless"):
                name = clean(raw.replace(",", ""))
            # a few items share a line with a staple or another item
            for extra in re.findall(r"\b(tuna|pistachios)\b", raw.lower()):
                canon, sec = canonical(extra)
                grouped[sec].add(canon)
            if not name or STAPLES.search(name):
                continue
            canon, sec = canonical(name)
            if canon:
                grouped[sec].add(canon)
                continue
            name = re.sub(r"^(cooked|chopped|grated|shredded|crumbled|toasted)\s+", "", name)
            for sec, pat in SECTIONS:
                if re.search(pat, name.lower()):
                    grouped[sec].add(name)
                    break
            else:
                grouped["Pantry"].add(name)
    return grouped


# what a prep step makes ahead, recognised from its wording
PREP_KEYS = {
    "MUFF": r"egg muffins", "OATS": r"overnight oats", "SOUP": r"Tuscan",
    "LENT": r"lentil salad", "TABB": r"tabbouleh", "DAL": r"\bdal\b",
    "CHICK": r"Chicken Grain Bowls", "MISO": r"miso tofu", "CONGEE": r"congee",
    "POZ": r"pozole", "CHILI": r"chili\b", "BOLO": r"Bolognese", "SOBA": r"soba",
    "PEPPER": r"stuff the peppers", "CHIA": r"chia pudding",
    "EGGS": r"boil (a batch of )?eggs|hard-boil eggs", "BITES": r"energy bites",
}
def code_ok(code, prepped):
    """Snacks only matter here when they are made ahead in the prep plan."""
    return code not in prepped or KEEPS[CODES[code]] is None


COOKED = {}   # code -> plan day (0 = Sunday before week 1) it was last cooked


def storage_notes(num, days, prep, snacks):
    """Check every make-ahead meal and leftover against its recipe's fridge
    life, and return the extra prep steps that keep each one in time:
    freezing a portion when the recipe freezes, a fresh batch when it doesn't."""
    sunday = (num - 1) * 7
    prepped = set()
    for code, pat in PREP_KEYS.items():
        step = next((x for x in prep if re.search(pat, x, re.I)), None)
        if step and "Saturday" not in step:
            COOKED[code] = sunday
            prepped.add(code)
    late, notes, midweek = {}, [], []
    for d, (meals, extra) in enumerate(zip(days, snacks), 1):
        today = sunday + d
        for slot, c in zip("BLD" + "S" * len(extra), list(meals) + extra):
            if slot == "S" and code_ok(c, prepped):
                continue
            if c == "FLEX":
                continue
            code = c.replace("LO:", "")
            name = CODES[code]
            keeps = KEEPS[name]
            batch = c.startswith("LO:") or (code in prepped and code in COOKED) or (
                slot != "D" and keeps and code in COOKED and COOKED[code] > sunday)
            if not batch or keeps is None:
                COOKED[code] = today
                continue
            if today - COOKED[code] <= keeps:
                continue
            if name in FREEZES:
                late.setdefault(code, []).append(DAYS[d - 1])
            else:
                COOKED[code] = today - 1
                midweek.append(f"On {DAYS[d - 2]} evening, make a fresh batch of the {name} "
                             f"for the rest of the week: it keeps {keeps} days in the fridge.")
    for code, when in late.items():
        name = CODES[code]
        notes.insert(0, f"Freeze the {' and '.join(when)} portion{'s' if len(when) > 1 else ''} "
                        f"of the {name} once cooled (it keeps {KEEPS[name]} days in the fridge), "
                        f"then move {'them' if len(when) > 1 else 'it'} to the fridge to thaw "
                        f"the night before.")
    return notes, midweek


def protein_of(code):
    if code == "FLEX":
        return FLEX_PROTEIN
    return PROTEIN[CODES[code.replace("LO:", "")]]


def label(code):
    if code == "FLEX":
        return "*Meal out (see tips)*"
    if code.startswith("LO:"):
        return f"Leftover {CODES[code[3:]]}"
    return CODES[code]


out = []
out.append("@chapter 18 | Weeks 1–12")
out.append("> Twelve weeks, eighty-four days, and not one evening spent wondering what's for dinner.")
out.append("")
out.append("Each week has a focus, a 7-day menu with snacks, a shopping list and a Sunday prep plan, followed by a short check-in. The figure beside each day shows its estimated protein total, including the planned snacks; every day lands between about 80 and 110 grams, in line with the 2025–2030 Dietary Guidelines range of roughly 82 to 109 grams for a woman of around 150 pounds. Leftovers are planned in, so most dinners also become a lunch, and every one is eaten within the fridge time given in its recipe; where a batch has to last longer, the prep plan tells you what to freeze or make fresh midweek.")
out.append("")
out.append("Quantities are in the recipes. Before you shop, check your pantry staples: olive oil, salt and pepper, spices, garlic, stock, canned tomatoes, oats and grains. And remember, this is a menu to borrow from, not a test: swap any meal for another with a similar protein figure, and skip whatever doesn't suit you.")
out.append("")
# the Printables PDF gets its own layout: one week per US Letter / A4 page,
# with the menu as a Day | Breakfast | Lunch | Dinner | Snack table
pr = ["@front The 12-Week Meal Plan",
      "One week per page: the menu with snacks and each day's estimated protein, "
      "the shopping list, the Sunday prep plan and a short check-in. "
      "Quantities and methods are in the book.", ""]
all_totals = []
for num, theme, focus, snacks, days, prep in WEEKS:
    out.append("@newpage")
    pr.append("@newpage")
    pr.append(f"## Week {num}: {theme}")
    pr.append("")
    pr.append("| Day@1.0 | Breakfast | Lunch | Dinner | Snack |")
    pr.append("|---|---|---|---|---|")
    if num in PHASES:
        out.append(f"## {PHASES[num]}")
        out.append("")
    out.append(f"## Week {num}: {theme}")
    out.append("")
    out.append(focus)
    out.append("")
    out.append("| Day@0.95 | Meals |")
    out.append("|---|---|")
    used, snack_days = set(), []
    for d, (b, l, n) in enumerate(days):
        base = protein_of(b) + protein_of(l) + protein_of(n)
        rot = snacks[d % len(snacks):] + snacks[:d % len(snacks)]
        chosen, total = [], base
        for s in rot:
            if total >= 80 and chosen:
                break
            sp = PROTEIN[CODES[s]]
            if total + sp > 110:
                continue
            chosen.append(s)
            total += sp
        all_totals.append(total)
        for c in (b, l, n):
            if c != "FLEX" and not c.startswith("LO:"):
                used.add(c)
        used.update(chosen)
        snack_days.append(chosen)
        snack_txt = " + ".join(CODES[s] for s in chosen) or "Fruit, if you're hungry"
        pr.append(f"| **{DAYS[d]}** <br> ≈{total} g protein | {label(b)} | {label(l)} | {label(n)} | {snack_txt} |")
        out.append(f"| {DAYS[d]} <br> ≈{total} g protein | **Breakfast:** {label(b)} <br> **Lunch:** {label(l)} <br> **Dinner:** {label(n)} <br> **Snack:** {snack_txt} |")
    out.append("")
    if any("FLEX" in day for day in days):
        out.append("**Eating out:** choose a dish built around fish, chicken, beans, eggs or tofu, add vegetables or a salad, and enjoy it. A meal out is counted at about 25 grams of protein.")
        out.append("")
    pr.append("")
    if any("FLEX" in day for day in days):
        pr.append("**Eating out:** build your plate around fish, chicken, beans, eggs or tofu, plus vegetables.")
        pr.append("")
    out.append("### Shopping list")
    out.append("")
    pr.append("### Shopping list")
    pr.append("")
    g = grocery(sorted(used))
    for sec in ["Produce", "Protein", "Dairy and alternatives", "Pantry", "Freezer"]:
        if g.get(sec):
            out.append(f"**{sec}:** " + "; ".join(sorted(set(g[sec]), key=str.lower)) + ".")
            out.append("")
            pr.append(f"[ ] **{sec}:** " + "; ".join(sorted(set(g[sec]), key=str.lower)) + ".")
    pr.append("")
    # Week 10 cooks three big batches, so its prep takes longer
    out.append("### Sunday prep (about 2½ hours)" if num == 10 else "### Sunday prep (60 to 90 minutes)")
    out.append("")
    freeze, midweek = storage_notes(num, days, prep, snack_days)
    pr.append("### Sunday prep (about 2½ hours)" if num == 10 else "### Sunday prep (60 to 90 minutes)")
    pr.append("")
    for i, step in enumerate(prep + freeze, 1):
        out.append(f"{i}. {step}")
        pr.append(f"{i}. {step}")
    out.append("")
    pr.append("")
    if midweek:
        out.append("**Midweek top-up:** " + " ".join(midweek))
        out.append("")
        pr.append("**Midweek top-up:** " + " ".join(midweek))
        pr.append("")
    pr.append("**Check-in:** Energy (1–10): ____   Sleep: ____   Symptoms that changed: ______________   "
              "Meals to keep: ______________   One win: ______________")
    pr.append("")
    out.append(":::try End-of-week check-in")
    out.append("[ ] How was my energy this week, from 1 to 10?")
    out.append("[ ] How did I sleep, and did anything help?")
    out.append("[ ] Which symptoms changed, for better or worse?")
    out.append("[ ] Which meals do I want to keep, and which can go?")
    out.append("[ ] One win I'm proud of this week:")
    out.append(":::")
    out.append("")

out.append(":::summary")
out.append("- Twelve themed weeks in three phases: Foundations, Build and Thrive.")
out.append("- Every day reaches roughly 80 to 110 grams of protein from three meals and planned snacks.")
out.append("- Leftovers are built in, so most dinners become a lunch, and a Sunday prep hour keeps the week easy.")
out.append("- The plan is a guide: swap meals freely, keep what works, and use Appendix B to build your own.")
out.append(":::")
out.append("")
out.append(">> Food is half of the plan. The other half happens on the floor of your living room, in thirty days of yoga, strength and calm.")
open(os.path.join(MS, "52_ch18.md"), "w", encoding="utf-8").write("\n".join(out) + "\n")
if os.environ.get("PRINT_OUT"):
    open(os.environ["PRINT_OUT"], "w", encoding="utf-8").write("\n".join(pr) + "\n")
print("wrote 52_ch18.md; daily protein range", min(all_totals), "-", max(all_totals))
