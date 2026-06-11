import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter

FONT = "Arial"
SLATE = "1F2A37"
HEADERBG = "374151"
BANDBG = "E5E7EB"
SUMBG = "F3F4F6"
CUSTOMBG = "FAFAFA"
WHITE = "FFFFFF"
GREY = "9CA3AF"

thin = Side(style="thin", color="D1D5DB")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

def cell(ws, coord, value=None, *, bold=False, size=10, color="111827", bg=None,
         italic=False, wrap=False, halign="left", valign="center", border=False):
    c = ws[coord]
    if value is not None:
        c.value = value
    c.font = Font(name=FONT, size=size, bold=bold, italic=italic, color=color)
    c.alignment = Alignment(horizontal=halign, vertical=valign, wrap_text=wrap)
    if bg:
        c.fill = PatternFill("solid", fgColor=bg)
    if border:
        c.border = BORDER
    return c

# ---- Standard 6S core (area-neutral) ----
CORE = [
    ("SORT", [
        ("Necessary Items Only", "Is the area free of obsolete, expired, surplus, damaged, or unidentifiable materials, parts, and product?"),
        ("Personal Belongings", "Are personal belongings (clothing, bags, appliances, decorative items) absent from the work area?"),
        ("Non-Area Items", "Is the area free of items that don't belong — scrap in wrong zones, furniture, unrelated equipment?"),
        ("Hazardous / Special Items", "Are combustible equipment, unsecured gas cylinders, and damaged materials removed or properly designated?"),
    ]),
    ("SET IN ORDER", [
        ("Location Labels", "Do all storage locations (shelves, bins, rack bays, zones) have clear, consistent, readable labels?"),
        ("Items in Correct Location", "Are all items stored in their labeled designated location — nothing ad-hoc or on the floor without reason?"),
        ("Aisle & Floor Clearance", "Are all aisles, pedestrian lanes, and floor areas clear — nothing blocking traffic flow or emergency egress?"),
        ("Equipment Parking", "Does all material-handling equipment (pallet jack, hand truck, ladders) have a designated home and is it stored there?"),
        ("Zone Definition", "Are distinct operational zones (staging, receiving, shipping, storage) clearly defined and marked?"),
    ]),
    ("SHINE", [
        ("Floor Cleanliness", "Are all floors clean and free of debris, dust, shavings, scraps, cardboard, and staining?"),
        ("Surfaces & Equipment", "Are shelving, work surfaces, bins, machines, and rack uprights clean and free of grime, dust, and oil?"),
        ("Waste & Cardboard Disposal", "Is waste disposed of promptly — bins not overflowing, cardboard broken down, packaging cleared daily?"),
        ("Facility Condition", "Are facilities in good condition — no damaged shelving, burned-out lights, water stains, or exposed wiring?"),
    ]),
    ("STANDARDIZE", [
        ("RACI Posted", "Is a current 6S RACI chart clearly posted in the area, identifying the responsible party for each S-category?"),
        ("Visual Layout / Zone Map", "Is a visual layout or zone map posted showing the correct 6S organization?"),
        ("Min / Max Inventory Levels", "Are minimum and maximum stock levels indicated for key inventory where applicable?"),
        ("Label Consistency", "Is the labeling system consistent — matching format, font, and location-coding standard?"),
        ("SOPs & Work References", "Are SOPs, machine instructions, and work references posted at the point of use?"),
    ]),
    ("SUSTAIN", [
        ("Audit History", "Was a 6S audit conducted in this area within the past 2 months with results posted in the area?"),
        ("Prior Issues Resolved", "Have all items flagged in the previous audit been addressed and verified corrected?"),
        ("Standard Maintained", "Does the area look as organized as the posted visual standard?"),
        ("Staff Awareness", "Can staff explain the 6S system, their RACI responsibilities, and locate items without assistance?"),
    ]),
    ("SAFETY", [
        ("Egress & Emergency Access", "Are all emergency exits, aisles, fire extinguishers, and electrical panels completely clear and unobstructed?"),
        ("Safe Storage & Stacking", "Are all items stored and stacked safely — no unstable loads, unsecured cylinders, or overhanging product?"),
        ("Hazard Controls", "Are machine safety zones marked, PPE requirements posted, and hazard controls in place?"),
        ("Fire Safety Equipment", "Is a fire extinguisher present, inspected within the past year, and unobstructed?"),
    ]),
]

CUSTOM_PER_CAT = 2

def build_room_form(ws, examples=None):
    """examples: dict {CATEGORY: [check_item, ...]} to pre-fill custom rows."""
    examples = examples or {}
    ws.sheet_view.showGridLines = False
    widths = {"A": 5, "B": 24, "C": 64, "D": 11, "E": 40}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w

    # Title
    ws.merge_cells("A1:E1")
    cell(ws, "A1", "6S + SAFETY AUDIT  ·  WORKSHEET", bold=True, size=15,
         color=WHITE, bg=SLATE, halign="center")
    ws.row_dimensions[1].height = 26
    # Legend
    ws.merge_cells("A2:E2")
    cell(ws, "A2", "0 = Major Issue      1 = Needs Work      2 = Good      3 = Excellent      "
                   "|      Comment required for any score below 2", size=9, italic=True,
         color="374151", bg=SUMBG, halign="center")
    ws.row_dimensions[2].height = 18

    # Field block
    cell(ws, "A3", "Location:", bold=True, size=10, halign="right")
    ws.merge_cells("B3:C3")
    cell(ws, "B3", "='Start Here'!$B$3", size=10, bg=WHITE, border=True)
    cell(ws, "D3", "Date:", bold=True, size=10, halign="right")
    cell(ws, "E3", None, bg=WHITE, border=True)
    cell(ws, "A4", "Room / Area:", bold=True, size=10, halign="right")
    ws.merge_cells("B4:C4")
    cell(ws, "B4", None, bg="FEF9C3", border=True)  # yellow = fill me in
    cell(ws, "D4", "Auditor:", bold=True, size=10, halign="right")
    cell(ws, "E4", None, bg=WHITE, border=True)

    # Summary band
    labels = ["Total Score", "Max Possible", "Score %", "# of Zeros", "RESULT"]
    for i, lab in enumerate(labels):
        cl = get_column_letter(1 + i)
        cell(ws, f"{cl}5", lab, bold=True, size=9, color=WHITE, bg=HEADERBG,
             halign="center", border=True)
    ws.row_dimensions[5].height = 16

    # Column headers (row 7); summary formulas filled after we know last row (row 6)
    hdr = ["No.", "Check Item", "Description / Audit Question", "Total Score\n(0–3)",
           "Comments  (required if score < 2)"]
    for i, h in enumerate(hdr):
        cl = get_column_letter(1 + i)
        cell(ws, f"{cl}7", h, bold=True, size=9, color=WHITE, bg=HEADERBG,
             halign="center", wrap=True, border=True)
    ws.row_dimensions[7].height = 28

    dv = DataValidation(type="whole", operator="between", formula1="0", formula2="3",
                        allow_blank=True, showErrorMessage=True)
    dv.error = "Enter a whole number 0–3."
    dv.errorTitle = "Score 0–3"
    ws.add_data_validation(dv)

    r = 8
    item_no = 1
    score_cells = []
    for cat, items in CORE:
        # band
        ws.merge_cells(f"A{r}:E{r}")
        cell(ws, f"A{r}", f"◆  {cat}", bold=True, size=10, color="111827", bg=BANDBG)
        ws.row_dimensions[r].height = 18
        r += 1
        ex = list(examples.get(cat, []))
        for name, desc in items:
            cell(ws, f"A{r}", f"{item_no}.0", size=9, halign="center", border=True)
            cell(ws, f"B{r}", name, size=9, wrap=True, valign="top", border=True)
            cell(ws, f"C{r}", desc, size=9, wrap=True, valign="top", border=True)
            cell(ws, f"D{r}", None, bg=WHITE, halign="center", border=True)
            cell(ws, f"E{r}", None, bg=WHITE, wrap=True, valign="top", border=True)
            dv.add(ws[f"D{r}"])
            score_cells.append(r)
            ws.row_dimensions[r].height = 30
            item_no += 1
            r += 1
        # custom rows
        for k in range(CUSTOM_PER_CAT):
            label = ex[k] if k < len(ex) else None
            cell(ws, f"A{r}", f"{item_no}.0", size=9, halign="center", color=GREY, bg=CUSTOMBG, border=True)
            cell(ws, f"B{r}", label, size=9, italic=(label is None), color=("111827" if label else GREY),
                 wrap=True, valign="top", bg=CUSTOMBG, border=True)
            cell(ws, f"C{r}", (None if label else "↳  Add a room-specific item for this category (optional)"),
                 size=9, italic=True, color=GREY, wrap=True, valign="top", bg=CUSTOMBG, border=True)
            cell(ws, f"D{r}", None, bg=WHITE, halign="center", border=True)
            cell(ws, f"E{r}", None, bg=WHITE, wrap=True, valign="top", border=True)
            dv.add(ws[f"D{r}"])
            score_cells.append(r)
            ws.row_dimensions[r].height = 28
            item_no += 1
            r += 1

    first, last = score_cells[0], score_cells[-1]
    rng = f"D{first}:D{last}"
    cnt = f"COUNT({rng})"
    tot = f"SUM({rng})"
    pct = f"{tot}/({cnt}*3)*100"
    zeros = f"COUNTIF({rng},0)"
    cell(ws, "A6", f"={tot}", bold=True, size=12, halign="center", bg=SUMBG, border=True)
    cell(ws, "B6", f"={cnt}*3", bold=True, size=12, halign="center", bg=SUMBG, border=True)
    cell(ws, "C6", f'=IF({cnt}=0,"",ROUND({pct},1))', bold=True, size=12, halign="center", bg=SUMBG, border=True)
    cell(ws, "D6", f"={zeros}", bold=True, size=12, halign="center", bg=SUMBG, border=True)
    cell(ws, "E6", f'=IF({cnt}=0,"",IF(AND({pct}>=67,{zeros}=0),"✔ PASS","✘ FAIL"))',
         bold=True, size=12, halign="center", bg=SUMBG, border=True)
    cell(ws, "A5", "Total Score", bold=True, size=9, color=WHITE, bg=HEADERBG, halign="center", border=True)

    # Page setup
    ws.print_area = f"A1:E{last}"
    ws.print_title_rows = "1:7"
    ws.page_setup.orientation = "portrait"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins.left = ws.page_margins.right = 0.4
    ws.page_margins.top = ws.page_margins.bottom = 0.5
    ws.oddFooter.right.text = "Page &P of &N"
    ws.oddFooter.left.text = "6S + Safety Audit"


wb = openpyxl.Workbook()

# ===== Start Here =====
ws = wb.active
ws.title = "Start Here"
ws.sheet_view.showGridLines = False
ws.column_dimensions["A"].width = 3
ws.column_dimensions["B"].width = 30
ws.column_dimensions["C"].width = 78
ws.merge_cells("B1:C1")
cell(ws, "B1", "6S PROGRAM — LOCATION AUDIT TEMPLATE", bold=True, size=16, color=WHITE, bg=SLATE, halign="center")
ws.row_dimensions[1].height = 30
cell(ws, "B3", "Location:", bold=True, size=12, halign="right")
cell(ws, "C3", None, bg="FEF9C3", border=True, size=12)
cell(ws, "C3").value = "[ Type your location / facility name here ]"
cell(ws, "C3", italic=True, color=GREY, bg="FEF9C3", border=True, size=11)
cell(ws, "B4", "Warehouse Lead:", bold=True, size=11, halign="right")
cell(ws, "C4", None, bg=WHITE, border=True)
cell(ws, "B5", "Audit Period / Year:", bold=True, size=11, halign="right")
cell(ws, "C5", None, bg=WHITE, border=True)

steps = [
    ("HOW TO USE THIS TEMPLATE", None),
    ("1. Set your location", "Type your facility name in the yellow Location cell above. It flows automatically into every room form."),
    ("2. Add a room", "Right-click the ROOM TEMPLATE tab → Duplicate. Rename the copy to the room (e.g. “Parts Room”). Repeat for every area you audit."),
    ("3. Name the room", "On each room form, type the room name in the yellow “Room / Area” cell."),
    ("4. Add custom checks", "Each form has the 26 standard 6S items plus blank custom rows under every category — add room-specific checks there (e.g. refrigerant cylinder storage, rack bay labels)."),
    ("5. Print blank forms (PDF)", "File → Print → set “Current sheet”, Portrait, Scale = Fit to width → Next → Save as PDF. Print and fill by hand during the walk-through."),
    ("6. Score & roll up", "Scores can be filled by hand on paper, or typed into the Score column to auto-calculate Total / % / PASS-FAIL. Record each room's result on the Facility Summary tab."),
    ("SCORING", None),
    ("Scale", "0 = Major Issue   |   1 = Needs Work   |   2 = Good   |   3 = Excellent. Comment required for any score below 2."),
    ("Pass rule", "Score ≥ 67% of the maximum AND no zeros on any item. Any Safety item scored 0 must be escalated to management the same day."),
    ("Frequency", "Audit 1–2 rooms per week so every room is graded monthly. Run a facility-wide review monthly or quarterly."),
]
r = 7
for head, body in steps:
    if body is None:
        ws.merge_cells(f"B{r}:C{r}")
        cell(ws, f"B{r}", head, bold=True, size=12, color=WHITE, bg=HEADERBG)
        ws.row_dimensions[r].height = 20
    else:
        cell(ws, f"B{r}", head, bold=True, size=10, valign="top", wrap=True)
        cell(ws, f"C{r}", body, size=10, valign="top", wrap=True)
        ws.row_dimensions[r].height = 32
    r += 1
ws.print_area = f"A1:C{r-1}"
ws.page_setup.orientation = "portrait"
ws.page_setup.fitToWidth = 1
ws.page_setup.fitToHeight = 0
ws.sheet_properties.pageSetUpPr.fitToPage = True

# ===== 6S Overview =====
ws = wb.create_sheet("6S Overview")
ws.sheet_view.showGridLines = False
ws.column_dimensions["A"].width = 3
ws.column_dimensions["B"].width = 22
ws.column_dimensions["C"].width = 92
ws.merge_cells("B1:C1")
cell(ws, "B1", "THE SIX S's", bold=True, size=16, color=WHITE, bg=SLATE, halign="center")
ws.row_dimensions[1].height = 28
elements = [
    ("Sort", "Go through everything in the area and keep only what you actually use. Remove broken tools, expired supplies, and clutter."),
    ("Set in Order", "Give everything a clear, logical home. The things you use most should be easiest to reach. Label locations so anyone can find — and return — an item."),
    ("Shine", "Clean the area and the equipment. A clean station makes it obvious when something is missing, leaking, or out of place."),
    ("Standardize", "Make the setup consistent. Use labels, shadow outlines, signs, or a simple checklist so the “right” way is obvious and repeatable."),
    ("Sustain", "Keep it up. Quick daily resets and the occasional check-in keep the area from sliding back."),
    ("Safety", "Build safety into the setup. Keep walkways and exits clear, store heavy or sharp items properly, keep PPE stocked, and fix or flag any hazard immediately."),
]
r = 3
for i, (name, desc) in enumerate(elements, 1):
    cell(ws, f"B{r}", f"{i}.  {name}", bold=True, size=11, valign="top", bg=BANDBG, border=True)
    cell(ws, f"C{r}", desc, size=10, valign="top", wrap=True, border=True)
    ws.row_dimensions[r].height = 40
    r += 1
ws.print_area = f"A1:C{r-1}"
ws.page_setup.orientation = "portrait"
ws.page_setup.fitToWidth = 1
ws.sheet_properties.pageSetUpPr.fitToPage = True

# ===== ROOM TEMPLATE =====
build_room_form(wb.create_sheet("ROOM TEMPLATE"))

# ===== Example - Warehouse =====
build_room_form(wb.create_sheet("Example - Warehouse"), examples={
    "SET IN ORDER": ["Rack Bay Labels", "Lineset & Flex Duct Storage"],
    "SHINE": ["Dock / Door Area"],
    "SAFETY": ["Refrigerant Cylinder Compliance", "Forklift / Pedestrian Lane Separation"],
})

# ===== Facility Summary =====
ws = wb.create_sheet("Facility Summary")
ws.sheet_view.showGridLines = False
for col, w in {"A": 28, "B": 16, "C": 12, "D": 14, "E": 44}.items():
    ws.column_dimensions[col].width = w
ws.merge_cells("A1:E1")
cell(ws, "A1", "FACILITY 6S SUMMARY", bold=True, size=15, color=WHITE, bg=SLATE, halign="center")
ws.row_dimensions[1].height = 26
cell(ws, "A2", "Location:", bold=True, size=10, halign="right")
cell(ws, "B2", "='Start Here'!$B$3", size=10, border=True)
cell(ws, "D2", "Audit Period:", bold=True, size=10, halign="right")
cell(ws, "E2", "='Start Here'!$C$5", size=10, border=True)
hdr = ["Area / Room", "Date Audited", "Score %", "Result", "Notes / Open Items"]
for i, h in enumerate(hdr):
    cl = get_column_letter(1 + i)
    cell(ws, f"{cl}4", h, bold=True, size=10, color=WHITE, bg=HEADERBG, halign="center", border=True)
for rr in range(5, 21):
    for i in range(5):
        cl = get_column_letter(1 + i)
        cell(ws, f"{cl}{rr}", None, bg=WHITE, border=True)
    ws.row_dimensions[rr].height = 22
cell(ws, "A22", "Facility Pass = every audited room scores ≥ 67% with no zeros.", italic=True, size=9, color="374151")
ws.print_area = "A1:E22"
ws.page_setup.orientation = "landscape"
ws.page_setup.fitToWidth = 1
ws.sheet_properties.pageSetUpPr.fitToPage = True

out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "6S-Audit-Template.xlsx")
wb.save(out)
print("saved", out)
