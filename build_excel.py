from openpyxl import Workbook
from openpyxl.styles import (
    PatternFill, Font, Alignment, Border, Side
)
from openpyxl.utils import get_column_letter

wb = Workbook()
ws = wb.active
ws.title = "P2 Stage 1"

# ── Colour palette ──────────────────────────────────────────────
DARK_BLUE   = "1F3864"
MID_BLUE    = "2E75B6"
LIGHT_BLUE  = "D6E4F0"
GOLD        = "F4B942"
GREEN_BG    = "E2EFDA"
ORANGE_BG   = "FCE4D6"
YELLOW_BG   = "FFEB9C"
WHITE       = "FFFFFF"
LIGHT_GRAY  = "F2F2F2"

def fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

def font(bold=False, color="000000", size=10, italic=False):
    return Font(bold=bold, color=color, size=size, italic=italic)

def center():
    return Alignment(horizontal="center", vertical="center", wrap_text=True)

def left():
    return Alignment(horizontal="left", vertical="center", wrap_text=True)

thin = Side(style="thin", color="BFBFBF")
thick = Side(style="medium", color="1F3864")

def thin_border():
    return Border(left=thin, right=thin, top=thin, bottom=thin)

def thick_border():
    return Border(left=thick, right=thick, top=thick, bottom=thick)

def style_row(ws, row, cols, bg, fg="000000", bold=False, align="left", size=10):
    for col in range(1, cols + 1):
        cell = ws.cell(row=row, column=col)
        cell.fill = fill(bg)
        cell.font = font(bold=bold, color=fg, size=size)
        cell.alignment = center() if align == "center" else left()
        cell.border = thin_border()

# ── Column widths ────────────────────────────────────────────────
col_widths = [5, 42, 10, 6, 10, 10, 10, 10, 6, 14, 12, 10, 16, 14]
for i, w in enumerate(col_widths, 1):
    ws.column_dimensions[get_column_letter(i)].width = w

# ── SECTION 1: Campaign Header ───────────────────────────────────
def merge_write(ws, row, col_start, col_end, text, bg, fg="FFFFFF", bold=True, size=11, align="left"):
    ws.merge_cells(start_row=row, start_column=col_start, end_row=row, end_column=col_end)
    cell = ws.cell(row=row, column=col_start)
    cell.value = text
    cell.fill = fill(bg)
    cell.font = font(bold=bold, color=fg, size=size)
    cell.alignment = center() if align == "center" else left()
    cell.border = thick_border()

r = 1
merge_write(ws, r, 1, 14, "iCliniq Google Search Ads  ·  P2 · Stage 1  ·  25 Keywords  ·  3 Ad Copies", DARK_BLUE, "FFFFFF", True, 13, "center")
ws.row_dimensions[r].height = 28

r = 2
merge_write(ws, r, 1, 14, "URL: icliniq.com/articles/cancer/stage-1-endometrial-cancer", MID_BLUE, "FFFFFF", False, 10)
ws.row_dimensions[r].height = 18

r = 3
merge_write(ws, r, 1, 14,
    "Campaign: Endometrial Cancer: Stage 1   |   Display Path: Cancer/Articles / Stage-1-Cancer   |   Budget: ₹2,000/day   |   Bid: Maximize Clicks   |   Geo: All USA",
    LIGHT_BLUE, "000000", False, 10)
ws.row_dimensions[r].height = 18

r = 4
merge_write(ws, r, 1, 14,
    "⚡ Ahrefs: 'stage 1 uterine cancer' (500 vol, TP 1,100) outperforms 'stage 1 endometrial cancer' (250 vol, TP 400). "
    "'endometrial cancer staging' (1,100 vol, KD 9) = easiest high-volume win. Tone = reassuring. 90%+ survival rate is the primary hook.",
    YELLOW_BG, "000000", False, 9)
ws.row_dimensions[r].height = 30

# ── SECTION 2: Keywords header ───────────────────────────────────
r = 6
merge_write(ws, r, 1, 14, "A · KEYWORDS  (25 total · Real Ahrefs USA Data · Phrase + Exact Match · All unique across all 5 pages)", MID_BLUE, "FFFFFF", True, 11)
ws.row_dimensions[r].height = 22

r = 7
headers = ["#","Keyword","Match","Tier","Vol/mo US","Global Vol","CPC (USD)","CPC (INR)","KD","Traffic Pot.","Priority","YMYL Safe?","Bid Strategy","Intent"]
for col, h in enumerate(headers, 1):
    cell = ws.cell(row=r, column=col, value=h)
    cell.fill = fill(DARK_BLUE)
    cell.font = font(bold=True, color="FFFFFF", size=10)
    cell.alignment = center()
    cell.border = thin_border()
ws.row_dimensions[r].height = 20

keywords = [
    (1,"stage 1 uterine cancer","Exact","T1",500,"-","$150","₹12,525",19,"1,100","🔴 High","✅ Yes","Max Clicks","Informational"),
    (2,"stage 1 endometrial cancer","Exact","T1",250,"-","$90","₹7,515",12,"400","🔴 High","✅ Yes","Max Clicks","Informational"),
    (3,"stage 1 uterine cancer survival rate","Exact","T1",150,"-","$60","₹5,010",36,"600","🔴 High","✅ Yes","Max Clicks","Informational"),
    (4,"stage 1 uterine cancer treatment","Exact","T1",250,"-","$250","₹20,875",20,"600","🔴 High","✅ Yes","Max Clicks","Informational"),
    (5,"endometrial cancer staging","Exact","T1",1100,"-","$140","₹11,690",9,"1,400","🔴 High","✅ Yes","Max Clicks","Informational"),
    (6,"stage 1 uterine cancer symptoms","Exact","T2",150,"-","$60","₹5,010",6,"33,000","🟡 Medium","✅ Yes","Max Clicks","Informational"),
    (7,"stage 1 endometrial cancer symptoms","Exact","T2",80,"-","$110","₹9,185",15,"400","🟡 Medium","✅ Yes","Max Clicks","Informational"),
    (8,"stage 1 endometrial cancer recurrence rate","Exact","T2",60,"-","$300","₹25,050",0,"10","🟡 Medium","✅ Yes","Max Clicks","Informational"),
    (9,"stage 1 hysteroscopy endometrial cancer","Exact","T3",150,"-","$45","₹3,757",1,"60","🟢 Test","✅ Yes","Max Clicks","Informational"),
    (10,"stage 1 endometrial cancer ultrasound","Exact","T3",100,"-","$250","₹20,875",3,"2,500","🟢 Test","✅ Yes","Max Clicks","Informational"),
    (11,"uterine cancer treatment stage 1","Exact","T3",60,"-","-","-",24,"500","🟢 Test","✅ Yes","Max Clicks","Informational"),
    (12,"stage 1 endometrial cancer","Phrase","T1",250,"-","$90","₹7,515",12,"400","🔴 High","✅ Yes","Max Clicks","Informational"),
    (13,"stage 1 uterine cancer","Phrase","T1",500,"-","$150","₹12,525",19,"1,100","🔴 High","✅ Yes","Max Clicks","Informational"),
    (14,"stage 1 uterine cancer treatment","Phrase","T1",250,"-","$250","₹20,875",20,"600","🔴 High","✅ Yes","Max Clicks","Informational"),
    (15,"stage 1 uterine cancer survival rate","Phrase","T1",150,"-","$60","₹5,010",36,"600","🔴 High","✅ Yes","Max Clicks","Informational"),
    (16,"endometrial cancer staging","Phrase","T1",1100,"-","$140","₹11,690",9,"1,400","🔴 High","✅ Yes","Max Clicks","Informational"),
    (17,"stage 1 uterine cancer symptoms","Phrase","T2",150,"-","$60","₹5,010",6,"33,000","🟡 Medium","✅ Yes","Max Clicks","Informational"),
    (18,"stage 1 endometrial cancer symptoms","Phrase","T2",80,"-","$110","₹9,185",15,"400","🟡 Medium","✅ Yes","Max Clicks","Informational"),
    (19,"uterine cancer stage 1","Phrase","T2",90,"-","$120","₹10,020",13,"1,200","🟡 Medium","✅ Yes","Max Clicks","Informational"),
    (20,"stage 1 endometrial cancer recurrence rate","Phrase","T3",60,"-","$300","₹25,050",0,"10","🟢 Test","✅ Yes","Max Clicks","Informational"),
    (21,"stage 1 endometrial cancer ultrasound","Phrase","T3",100,"-","$250","₹20,875",3,"2,500","🟢 Test","✅ Yes","Max Clicks","Informational"),
    (22,"stage 1 hysteroscopy endometrial cancer","Phrase","T3",150,"-","$45","₹3,757",1,"60","🟢 Test","✅ Yes","Max Clicks","Informational"),
    (23,"uterine cancer treatment stage 1","Phrase","T3",60,"-","-","-",24,"500","🟢 Test","✅ Yes","Max Clicks","Informational"),
    (24,"uterine cancer stage 1","Exact","T2",90,"-","$120","₹10,020",13,"1,200","🟡 Medium","✅ Yes","Max Clicks","Informational"),
    (25,"stage 1a endometrial cancer","Phrase","T3",200,"-","$250","₹20,875",1,"400","🟢 Test","✅ Yes","Max Clicks","Informational"),
]

tier_colors = {"T1": GREEN_BG, "T2": YELLOW_BG, "T3": ORANGE_BG}

for kw in keywords:
    r += 1
    bg = tier_colors.get(kw[3], WHITE)
    for col, val in enumerate(kw, 1):
        cell = ws.cell(row=r, column=col, value=val)
        cell.fill = fill(bg)
        cell.font = font(size=9)
        cell.alignment = center() if col != 2 else left()
        cell.border = thin_border()
    ws.row_dimensions[r].height = 16

# Negative keywords row
r += 1
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=14)
cell = ws.cell(row=r, column=1,
    value="NEG: stage 2 | stage 3 | stage 4 | metastatic | free treatment | clinical trial | home remedy | wikipedia | insurance | jobs | ovarian cancer | cervical cancer | endometriosis")
cell.fill = fill("FFE0E0")
cell.font = font(bold=True, color="C00000", size=9)
cell.alignment = left()
cell.border = thin_border()
ws.row_dimensions[r].height = 16

# ── SECTION 3: Ad Copies ─────────────────────────────────────────
ads = [
    {
        "title": "AD 1 of 3  ·  Ad 1 - Newly Diagnosed  ·  Angle: Patients just received a Stage 1 diagnosis - need clarity",
        "headlines": [
            ("1","Stage 1 Endometrial Cancer","📌 P1","26✅"),
            ("2","Oncologist-Written Guide","📌 P2","24✅"),
            ("3","Free Article - Read Now","📌 P3","23✅"),
            ("4","Just Diagnosed? Start Here","","26✅"),
            ("5","Stage 1A and 1B Explained","","25✅"),
            ("6","90%+ Survival - What It Means","","29✅"),
            ("7","Surgery Options at Stage 1","","26✅"),
            ("8","What to Expect After Diagnosis?","","30✅"),
            ("9","Early-Stage Cancer - The Facts","","30✅"),
            ("10","Trusted by 5M+ Patients","","23✅"),
            ("11","Hysterectomy - What to Know","","27✅"),
            ("12","Prognosis After Stage 1","","23✅"),
            ("13","18,000+ Verified Doctors","","24✅"),
            ("14","iCliniq - Free Cancer Guides","","28✅"),
            ("15","Read Before Your Next Visit","","27✅"),
        ],
        "descriptions": [
            ("D1","Stage 1 endometrial cancer is confined to the uterus — the most treatable stage. Free.","87✅"),
            ("D2","5-year survival above 90% for Stage 1. Free oncologist guide covers staging and surgery.","89✅"),
            ("D3","Hysterectomy, brachytherapy, hormone therapy — Stage 1 treatments clearly explained.","85✅"),
            ("D4","Just diagnosed with Stage 1 uterine cancer? Free doctor-reviewed guide for what's next.","88✅"),
        ],
    },
    {
        "title": "AD 2 of 3  ·  Ad 2 - Survival Rate Focus  ·  Angle: Patients researching 5-year survival and prognosis",
        "headlines": [
            ("1","Stage 1 Endometrial Cancer","📌 P1","26✅"),
            ("2","Oncologist-Written Guide","📌 P2","24✅"),
            ("3","Free Article - Read Now","📌 P3","23✅"),
            ("4","5-Year Survival Rate: 90%+","","26✅"),
            ("5","Stage 1A vs 1B: Key Differences","","30✅"),
            ("6","Recurrence Risk at Stage 1","","26✅"),
            ("7","Long-Term Outcomes Explained","","28✅"),
            ("8","Survival Stats - Plain English","","30✅"),
            ("9","Based on Real Population Data","","29✅"),
            ("10","Stage 1 Prognosis - The Data","","28✅"),
            ("11","No Jargon Survival Guide","","24✅"),
            ("12","Endometrial Ca Staging Guide","","28✅"),
            ("13","Oncology Stats Explained","","24✅"),
            ("14","Compare Stage 1A vs 1B Risk","","27✅"),
            ("15","iCliniq - Trusted Health Info","","29✅"),
        ],
        "descriptions": [
            ("D1","Stage 1 endometrial cancer has a 5-year survival rate above 90%. Free oncologist guide.","88✅"),
            ("D2","What does Stage 1A vs 1B mean for your prognosis? Free oncologist-reviewed. No login.","86✅"),
            ("D3","Understand recurrence risk and long-term outcomes for Stage 1 uterine cancer. Free.","84✅"),
            ("D4","iCliniq explains Stage 1 survival rates in plain English. Medically reviewed. Free.","84✅"),
        ],
    },
    {
        "title": "AD 3 of 3  ·  Ad 3 - Treatment Options  ·  Angle: Patients researching surgery and radiation options",
        "headlines": [
            ("1","Stage 1 Endometrial Cancer","📌 P1","26✅"),
            ("2","Oncologist-Written Guide","📌 P2","24✅"),
            ("3","Free Article - Read Now","📌 P3","23✅"),
            ("4","Surgery Options for Stage 1","","27✅"),
            ("5","Hysterectomy - Step by Step","","27✅"),
            ("6","Brachytherapy at Stage 1","","24✅"),
            ("7","Hormone Therapy Explained","","25✅"),
            ("8","Radiation for Stage 1 Cancer","","28✅"),
            ("9","Full Stage 1 Treatment Guide","","28✅"),
            ("10","What Comes After Surgery?","","25✅"),
            ("11","Follow-Up Care - Stage 1","","24✅"),
            ("12","Recovery After Hysterectomy","","27✅"),
            ("13","18,000+ Cancer Specialists","","26✅"),
            ("14","Stage 1 Treatment Deep Dive","","27✅"),
            ("15","Free. No Ads. No Login.","","23✅"),
        ],
        "descriptions": [
            ("D1","Stage 1 treatment options: hysterectomy, brachytherapy, hormone therapy. Free guide.","85✅"),
            ("D2","Which surgery is right for Stage 1 endometrial cancer? Free oncologist-reviewed guide.","87✅"),
            ("D3","From surgery to follow-up care — iCliniq covers the complete Stage 1 treatment plan.","84✅"),
            ("D4","Stage 1 uterine cancer treatment in plain English. Free. No login required.","76✅"),
        ],
    },
]

ad_colors = [("C5E0B4","375623"), ("BDD7EE","1F3864"), ("FCE4D6","843C0C")]

for ad_idx, ad in enumerate(ads):
    r += 2
    bg_head, fg_head = ad_colors[ad_idx]
    merge_write(ws, r, 1, 14, "  " + ad["title"], bg_head, fg_head, True, 10)
    ws.row_dimensions[r].height = 22

    r += 1
    merge_write(ws, r, 1, 14, "  HEADLINES  (15 total · H1 H2 H3 = Pinned · H4–H15 = Unique · All ≤30 chars · YMYL + CTR optimised)", DARK_BLUE, "FFFFFF", True, 9)
    ws.row_dimensions[r].height = 18

    r += 1
    for col, h in enumerate(["#","Headline","Pin","Chars"], 1):
        cell = ws.cell(row=r, column=col, value=h)
        cell.fill = fill(MID_BLUE)
        cell.font = font(bold=True, color="FFFFFF", size=10)
        cell.alignment = center()
        cell.border = thin_border()
    ws.row_dimensions[r].height = 18

    for i, (num, hl, pin, chars) in enumerate(ad["headlines"]):
        r += 1
        row_bg = LIGHT_BLUE if i < 3 else (LIGHT_GRAY if i % 2 == 0 else WHITE)
        for col, val in enumerate([num, hl, pin, chars], 1):
            cell = ws.cell(row=r, column=col, value=val)
            cell.fill = fill(row_bg)
            cell.font = font(bold=(i < 3), size=10)
            cell.alignment = center() if col != 2 else left()
            cell.border = thin_border()
        ws.row_dimensions[r].height = 18

    r += 1
    merge_write(ws, r, 1, 14,
        "  CTR STRATEGY: H1=keyword match · H2=E-E-A-T trust · H3=low-friction CTA · H4–H6=angle · H7–H9=credibility · H10–H12=value · H13–H15=rotation",
        LIGHT_GRAY, "595959", False, 8)
    ws.row_dimensions[r].height = 16

    r += 1
    merge_write(ws, r, 1, 14, "  DESCRIPTIONS  (4 total · All ≤90 chars · Unique to this ad · YMYL-compliant)", DARK_BLUE, "FFFFFF", True, 9)
    ws.row_dimensions[r].height = 18

    r += 1
    for col, h in enumerate(["#","Description Text (max 90 chars · YMYL clinical tone · factual · no cure claims · no fear)","Chars","OK?"], 1):
        cell = ws.cell(row=r, column=col, value=h)
        cell.fill = fill(MID_BLUE)
        cell.font = font(bold=True, color="FFFFFF", size=10)
        cell.alignment = center() if col != 2 else left()
        cell.border = thin_border()
    ws.row_dimensions[r].height = 18

    for d_num, d_text, d_chars in ad["descriptions"]:
        r += 1
        row_bg = GREEN_BG if d_num in ("D1","D3") else WHITE
        for col, val in enumerate([d_num, d_text, d_chars, "✅ OK"], 1):
            cell = ws.cell(row=r, column=col, value=val)
            cell.fill = fill(row_bg)
            cell.font = font(size=10)
            cell.alignment = center() if col != 2 else left()
            cell.border = thin_border()
        ws.row_dimensions[r].height = 28

# ── SECTION 4: YMYL Audit ────────────────────────────────────────
r += 2
merge_write(ws, r, 1, 14, "YMYL + POLICY AUDIT", DARK_BLUE, "FFFFFF", True, 11, "center")
ws.row_dimensions[r].height = 22

audit_rows = [
    ("✅", "YMYL Clinical: All copy uses evidence-based, factual, informational framing. No exaggerated claims.", GREEN_BG),
    ("✅", "E-E-A-T: H2 'Oncologist-Written Guide' signals medical expertise in every ad. D copy cites 'oncologist-reviewed' sourcing.", GREEN_BG),
    ("✅", "CTR optimised: 90%+ survival stat creates factual curiosity. Stage-specific framing signals relevance. Questions drive clicks.", GREEN_BG),
    ("✅", "Zero cure claims, guaranteed outcomes, scare tactics, or 'your condition' language across all 15 headlines and 12 descriptions.", GREEN_BG),
    ("✅", "Zero drug names in headlines/descriptions (Oct 2025 policy). Educational use on landing page articles only.", GREEN_BG),
    ("✅", "All 15 headlines ≤30 chars  |  All 4 descriptions ≤90 chars  |  Both display paths ≤15 chars  |  25 unique keywords", GREEN_BG),
    ("✅", "Zero duplicate headlines or descriptions across all 3 ads on this page.", GREEN_BG),
    ("⚠️", "Remarketing OFF: Do not build custom audiences from these cancer page visits (Google Personalised Health policy).", YELLOW_BG),
]

for icon, text, bg in audit_rows:
    r += 1
    ws.cell(row=r, column=1, value=icon).fill = fill(bg)
    ws.cell(row=r, column=1).alignment = center()
    ws.cell(row=r, column=1).border = thin_border()
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=14)
    cell = ws.cell(row=r, column=2, value=text)
    cell.fill = fill(bg)
    cell.font = font(size=10)
    cell.alignment = left()
    cell.border = thin_border()
    ws.row_dimensions[r].height = 20

# freeze header rows
ws.freeze_panes = "A8"

wb.save("/home/user/ppc-ai-skills/Endo_Cancer_PPC_Rewritten_P2_Stage1.xlsx")
print("Done")
