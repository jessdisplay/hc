"""Rebuild Harmony's compliance calendar in the Avaana checklist layout.

Brittany asked for one checklist in the shape of Avaana's 2023 NDIS application
checklist: module, Practice Standard, outcome, indicators, then the documents.
This takes the 16 Oct 2026 compliance calendar (every audited document, its link
and its review dates) and replaces its flat 'All documents' tab with a
'Checklist' tab in that shape, adding version columns and highlighting every
form someone must fill in by whose file it belongs in. The other tabs are kept.

It is a one-off: once Harmony types review dates into the result, that workbook
is the master, and export_web.py reads it for the website page.

    python3 make_workbook.py BASE.xlsx OUT.xlsx
"""
import json
import re
import sys
from pathlib import Path

import openpyxl
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from send_lists import SEND_TABS

HERE = Path(__file__).parent
STANDARDS = json.loads((HERE / "standards.json").read_text())

# Harmony Care Palette 01 (site/index.html tokens).
FOREST, SAGE, ROSE, CREAM, INK, MUTED, RULE = (
    "38553E", "9EAE8F", "F17990", "F3E3CA", "2E2A26", "5F564C", "DFD5C4")
TINT = {"Participant": "FBE0E6", "Worker": "DCE7D5", "Organisation": CREAM}  # rose and sage, lightened
FILE_LABEL = {"Participant": "Each participant", "Worker": "Each worker",
              "Organisation": "Organisation"}

MODULES = [  # (module_ord in the Guidelines, label)
    (1, "Core Module"),
    (2, "Module 1: High intensity daily personal activities"),
    (4, "Module 2A: Implementing behaviour support plans"),
    (6, "Module 4: Specialised support coordination"),
    (8, "Module 5A: Supported independent living"),
]
# Groups that are not a single Guidelines section.
EXTRA = {
    "M1-ALL": (2, "Every Module 1 support",
               "These documents serve every high intensity support Harmony delivers "
               "(bowel care, enteral feeding, urinary catheter and wound management)."),
    "M5A-ALL": (8, "Every SIL standard",
                "These documents serve every part of Module 5A."),
    "BUS": (9, "Business and registration evidence",
            "Not a Practice Standard. The papers the Commission and the auditor ask to see: "
            "business plan, insurances, registrations and the application answers."),
}

AREA_TO_SECTION = {
    "Rights: Person-centred supports": "6",
    "Rights: Privacy and dignity": "8",
    "Rights: Independence and informed choice": "9",
    "Rights: Violence, abuse, neglect, exploitation and discrimination": "10",
    "Governance: Governance and operational management": "11",
    "Governance: Risk management": "12",
    "Governance: Quality management": "13",
    "Governance: Information management": "14",
    "Governance: Feedback and complaints": "15",
    "Governance: Incident management": "16",
    "Governance: Human resource management": "17",
    "Governance: Continuity of supports": "18",
    "Governance: Emergency and disaster management": "18A",
    "Provision: Access to supports": "19",
    "Provision: Support planning": "20",
    "Provision: Service agreements": "21",
    "Provision: Responsive support provision": "22",
    "Provision: Transitions to or from the provider": "23",
    "Environment: Safe environment": "24",
    "Environment: Infection prevention and control": "24",  # IPC is indicators 7 to 11 of s 24
    "Environment: Client money and property": "25",
    "Environment: Management of medication": "26",
    "Environment: Mealtime management": "26A",
    "Environment: Management of waste": "27",
    "Business and registration": "BUS",
}
# The module areas hold several standards each, so place those documents by title.
# First match wins; every module document must match something (checked below).
TITLE_TO_SECTION = [
    ("Module 1", r"bowel", "29"),
    ("Module 1", r"enteral|tube", "30"),
    ("Module 1", r"catheter", "32"),
    ("Module 1", r"wound", "35"),
    ("Module 1", r".", "M1-ALL"),
    ("Module 2A", r"^Implementing Behaviour Supports Plans Policy|Internal Audit Schedule", "47"),
    ("Module 2A", r"Template Letter to PBSP", "49"),
    ("Module 2A", r"\(Interim\)", "54"),
    ("Module 2A", r"^Restrictive Practices Register", "51"),
    ("Module 2A", r"^Review Register|ABC Data", "52"),
    ("Module 2A", r".", "50"),
    ("Module 4", r"Conflict of Interest", "66"),
    ("Module 4", r"Services Agreement", "65"),
    ("Module 4", r".", "64"),
    ("Module 5a", r"Supported Decision-Making", "72B"),
    ("Module 5a", r"Safeguarding|Co-Tenancy Concern", "72C"),
    ("Module 5a", r"Practice Governance|Emergency Plan|Vision, Values|Training Plan|Induction to a House"
                  r"|Co-Tenant Consultation|House Meeting|Competency Assessment", "72D"),
    ("Module 5a", r"Tenancy|Transition or Exit|Services Agreement|Schedule to the Services|Collaboration"
                  r"|Conflict of Interest", "72E"),
    ("Module 5a", r".", "M5A-ALL"),
]
# Diversity Policy is filed under person-centred supports but is the policy for
# 'Individual values and beliefs', which otherwise has no document of its own.
TITLE_OVERRIDE = {"Diversity Policy": "7"}

FILLED_IN = {"Form", "Register", "Plan", "Record", "Checklist", "Chart", "Agreement", "Template"}
FILE_OF = {"Each client": "Participant", "Each behaviour support plan": "Participant",
           "Each worker": "Worker", "Each new worker": "Worker",
           "Organisation": "Organisation", "Each SIL house": "Organisation",
           "Each incident": "Organisation", "Each complaint": "Organisation"}
# A template the organisation fills in to make a document, not a worker's own form.
FILE_OVERRIDE = {"Position Description Template": "Organisation"}

# Harmony's Policy Register (edited 28 Aug 2026): register title -> our title.
POLICY_REGISTER = {
    "Choice Advocacy and Control Policy": "Choice, Advocacy and Control Policy",
    "Diversity Policy": "Diversity Policy",
    "Preventing and Responding to Violence, Abuse, Neglect, Exploitation and Discrimination Policy":
        "Preventing and Responding to Violence, Abuse, Neglect, Exploitation and Discrimination Policy",
    "Privacy Policy": "Privacy and Dignity Policy",
    "Promoting and Protecting Rights Policy": "Promoting and Protecting Rights Policy",
    "Feedback and Complaints Management Policy": "Feedback and Complaints Management Policy",
    "Incident Management and Reporting Policy": "Incident Management and Reporting Policy",
    "Quality Management and Continuous Improvement Policy": "Quality Management and Continuous Improvement Policy",
    "Risk Management Policy": "Risk Management Policy",
    "Emergency and Disaster Preparedness Policy": "Emergency and Disaster Preparedness Policy",
    "Work Health and Safety Policy": "Work Health and Safety Policy",
    "COVID-19 Pandemic Management Policy": "COVID-19 Pandemic Management Policy",
    "Governance Policy": "Governance Policy",
    "Human Resources Management Policy": "Human Resources Management Policy",
    "Information Management Policy": "Information Management Policy",
    "Conflict of Interest Policy": "Conflict of Interest Policy",
    "Provision of Supports Policy": "Provision of Supports Policy",
    "Client Money and Property Policy": "Client Money and Property Policy",
    "Infection Control Policy": "Infection Control Policy",
    "Support Coordination Policy": "Support Coordination Policy",
    "Supported Independent Living Policy": "Supported Independent Living Policy (Version 1, core folder copy)",
    "Mealtime Management Policy": "Mealtime Management Policy",
    "Management of Medication Policy": "Management of Medication Policy",
    "Waste Management Policy": "Waste Management Policy",
    "Implementing Behaviour Supports Plans Policy & Procedures":
        "Implementing Behaviour Supports Plans Policy and Procedures",
}
VERSION_KNOWN = {"Supported Independent Living Policy": "2"}  # replaces Version 1 (6 Oct 2026 pack)

COLS = [  # (header, width, key)
    ("No.", 6, "no"),
    ("Module", 16, "module"),
    ("Practice Standard", 24, "standard"),
    ("Document (opens the file)", 42, "title"),
    ("Type", 13, "type"),
    ("Filled in for", 14, "file"),
    ("When it is needed", 40, "when"),
    ("Completed by", 22, "completed_by"),
    ("Signed by", 26, "signed_by"),
    ("Version", 9, "version"),
    ("Approved by", 17, "approved_by"),
    ("Approval date", 13, "approval_date"),
    ("Next review on record", 13, "next_review"),
    ("Review every (months)", 10, "review_months"),
    ("Last reviewed (type the date)", 14, "last_reviewed"),
    ("Next due", 12, None),
    ("Status", 18, None),
    ("Notes for the next audit", 80, "notes"),
    ("How often", 40, "how_often"),
    ("Kept where", 24, "kept_where"),
    ("Harmony's rule (source)", 46, "rule"),
    ("Section", 8, "section"),  # hidden: the Guidelines section, read by export_web.py
]
COL = {key or head: i + 1 for i, (head, _, key) in enumerate(COLS)}
TYPED = {"version", "approved_by", "approval_date", "last_reviewed"}
HEADER_ROW = 4

thin = Side(style="thin", color=RULE)
BORDER = Border(bottom=thin)
WRAP = Alignment(wrap_text=True, vertical="top")
F = lambda size=10, **k: Font(name="Arial", size=size, **k)  # noqa: E731


def fill(rgb):
    return PatternFill("solid", fgColor=rgb)


def read_base(path):
    wb = openpyxl.load_workbook(path)
    ws = wb["All documents"]
    head = [c.value for c in ws[1]]
    docs = []
    for row in ws.iter_rows(min_row=2):
        if not row[1].value:
            continue
        d = dict(zip(head, (c.value for c in row)))
        status = d["Status"]
        docs.append({
            "area": d["Area"],
            "title": d["Document (opens the file)"],
            "url": row[1].hyperlink.target if row[1].hyperlink else None,
            "type": d["Type"], "for": d["For"],
            "when": d["When it is needed"], "how_often": d["How often"],
            "completed_by": d["Completed by"], "signed_by": d["Signed by"],
            "kept_where": d["Kept where"],
            "next_review": d["Next review on record"],
            "review_months": d["Review every (months)"],
            "last_reviewed": d["Last reviewed (type the date)"],
            "rule": d["Harmony's rule (source)"], "notes": d["Notes for the next audit"],
            # A few rows carry a fixed status (working copy, superseded) instead of the formula.
            "fixed_status": None if (status is None or str(status).startswith("=")) else status,
        })
    return wb, docs


def place(doc):
    if doc["title"] in TITLE_OVERRIDE:
        return TITLE_OVERRIDE[doc["title"]]
    if doc["area"] in AREA_TO_SECTION:
        return AREA_TO_SECTION[doc["area"]]
    area = doc["area"].split(":")[0]
    for a, pattern, section in TITLE_TO_SECTION:
        if a == area and re.search(pattern, doc["title"], re.I):
            return section
    raise SystemExit(f"no Practice Standard for {doc['area']!r} / {doc['title']!r}")


def enrich(docs):
    by_title = {d["title"]: d for d in docs}
    for reg_title, ours in POLICY_REGISTER.items():
        d = by_title.get(ours)
        if d is None:
            raise SystemExit(f"Policy Register entry {reg_title!r} matches no document ({ours!r})")
        d.update(version="1", approved_by="Victoria J Scilipote", approval_date="October 2023")
    for title, v in VERSION_KNOWN.items():
        by_title[title]["version"] = v
    for d in docs:
        d["section"] = place(d)
        if d["type"] in FILLED_IN:
            d["file"] = FILE_OVERRIDE.get(d["title"]) or FILE_OF[d["for"]]
        else:
            d["file"] = None


def groups(docs):
    """Yield (module label, section key, heading, outcome, rows, clause, docs) in Guidelines order."""
    sections = {s["section_no"]: s for s in STANDARDS["sections"]}
    order = []
    for ord_, label in MODULES + [(9, "Business and registration")]:
        keys = [k for k, (m, _, _) in EXTRA.items() if m == ord_ and k != "BUS"]
        keys += [s["section_no"] for s in STANDARDS["sections"] if s["module_ord"] == ord_]
        keys += ["BUS"] if ord_ == 9 else []
        for k in keys:
            mine = [d for d in docs if d["section"] == k]
            if not mine:
                continue
            if k in EXTRA:
                _, heading, outcome = EXTRA[k]
                order.append((label, k, heading, outcome, [], None, mine))
            else:
                s = sections[k]
                order.append((label, k, s["heading"], s["outcome"], s["rows"],
                              f"Practice Standards Rules, {label.split(':')[0]}, {s['clause_label'].lower()}; "
                              f"Quality Indicators Guidelines s {k}", mine))
    placed = sum(len(g[-1]) for g in order)
    if placed != len(docs):
        raise SystemExit(f"{len(docs) - placed} documents fell outside every group")
    return order


def indicator_lines(rows):
    """One string per top-level indicator, its sub-items on their own lines."""
    out = []
    for i, r in enumerate(rows):
        if r["parent"] is None:
            out.append([f"{r['item']}. {r['text']}"])
        else:
            out[-1].append(f"     ({r['item']}) {r['text']}")
        for c in r.get("continued") or []:
            out[-1].append(f"     {c}")
    return ["\n".join(x) for x in out]


def height(text, width_units, size=10):
    """Row height in points that fits wrapped text in a cell this many units wide."""
    per_line = max(1, int(width_units * 1.15 * 10 / size))
    lines = sum(max(1, -(-len(p) // per_line)) for p in str(text).split("\n"))
    return min(409, max(15, lines * size * 1.32 + 4))


def write_checklist(wb, order, at):
    ws = wb.create_sheet("Checklist", at)
    for i, (_, w, _) in enumerate(COLS, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.column_dimensions[get_column_letter(COL["section"])].hidden = True
    last = get_column_letter(len(COLS))
    span = sum(w for _, w, k in COLS[COL["title"] - 1:COL["signed_by"]])
    d0, d1 = get_column_letter(COL["title"]), get_column_letter(COL["signed_by"])

    ws["A1"] = "Harmony Care Support Services: documents and forms by Practice Standard"
    ws["A1"].font = F(bold=True, size=14, color=FOREST)
    ws["A2"] = "Highlighted rows are forms someone must fill in, coloured by whose file they belong in:"
    ws["A2"].font = F(color=MUTED)
    for col, who in zip(("F", "G", "H"), ("Participant", "Worker", "Organisation")):
        ws[f"{col}2"] = FILE_LABEL[who] if who != "Organisation" else "The organisation"
        ws[f"{col}2"].fill = fill(TINT[who])
        ws[f"{col}2"].font = F(bold=True)
        ws[f"{col}2"].alignment = Alignment(horizontal="center")
    ws["I2"] = "No colour: policies and guides, read but not filled in"
    ws["I2"].font = F(color=MUTED)
    ws["A3"] = ("Wording of each outcome and indicator: NDIS Practice Standards Rules (compilation 6) and "
                "Quality Indicators Guidelines (compilation 3), both in force 1 July 2026. The 1 and 2 "
                "buttons at the top left fold the indicators away and back.")
    ws["A3"].font = F(color=MUTED, italic=True)

    for i, (head, _, key) in enumerate(COLS, 1):
        c = ws.cell(HEADER_ROW, i, head)
        c.font = F(bold=True, color="FFFFFF")
        c.fill = fill(SAGE if key in TYPED else FOREST)
        if key in TYPED:
            c.font = F(bold=True, color=INK)
        c.alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[HEADER_ROW].height = 44

    r = HEADER_ROW + 1
    no = 0
    module_seen = None
    doc_rows = []
    for module, key, heading, outcome, rows, clause, docs in order:
        if module != module_seen:
            module_seen = module
            ws.merge_cells(f"A{r}:{d1}{r}")
            c = ws.cell(r, 1, module.upper())
            c.font = F(bold=True, size=12, color="FFFFFF")
            c.fill = fill(FOREST)
            c.alignment = Alignment(vertical="center")
            ws.row_dimensions[r].height = 22
            r += 1
        # The standard: its name, then its outcome across the document columns.
        ws.cell(r, COL["standard"], heading).font = F(bold=True, color=FOREST, size=11)
        ws.cell(r, COL["standard"]).alignment = WRAP
        ws.merge_cells(f"{d0}{r}:{d1}{r}")
        text = f"Outcome: {outcome}" + (f"\n({clause})" if clause else "")
        c = ws.cell(r, COL["title"], text)
        c.font = F(italic=True, color=INK)
        c.alignment = WRAP
        for i in range(1, len(COLS)):
            ws.cell(r, i).border = Border(top=Side(style="medium", color=FOREST))
        ws.row_dimensions[r].height = max(height(heading, 24, 11), height(text, span))
        r += 1
        lines = indicator_lines(rows)
        first = r
        for n, line in enumerate(lines):
            ws.cell(r, COL["standard"], "Indicators to be demonstrated" if n == 0 else None).font = F(color=MUTED, size=9)
            ws.merge_cells(f"{d0}{r}:{d1}{r}")
            c = ws.cell(r, COL["title"], line)
            c.font = F(color=MUTED, size=9)
            c.alignment = WRAP
            ws.row_dimensions[r].height = height(line, span, 9)
            r += 1
        if lines:
            ws.row_dimensions.group(first, r - 1, outline_level=1, hidden=False)
        for d in docs:
            no += 1
            d["no"] = no
            d["module"] = module.split(":")[0]
            d["standard"] = heading
            for i, (_, w, k) in enumerate(COLS, 1):
                if k is None:
                    continue
                v = d.get(k)
                c = ws.cell(r, i, v)
                c.font = F()
                c.alignment = WRAP
                c.border = BORDER
                if k == "file":
                    c.value = FILE_LABEL[v] if v else "Not filled in"
                    c.font = F(bold=bool(v), color=INK if v else MUTED)
                if k in ("next_review", "last_reviewed") and v:
                    c.number_format = "d mmm yyyy"
            t = ws.cell(r, COL["title"])
            if d["url"]:
                t.hyperlink = d["url"]
                t.font = F(color=FOREST, underline="single", bold=bool(d["file"]))
            L = lambda k: f"{get_column_letter(COL[k])}{r}"  # noqa: E731
            due = ws.cell(r, COL["Next due"],
                          f'=IF({L("last_reviewed")}<>"",IF({L("review_months")}="","",'
                          f'EDATE({L("last_reviewed")},{L("review_months")})),'
                          f'IF({L("next_review")}="","",{L("next_review")}))')
            due.number_format = "d mmm yyyy"
            due.font, due.border, due.alignment = F(), BORDER, WRAP
            st = ws.cell(r, COL["Status"], d["fixed_status"] or (
                f'=IF({L("Next due")}="","No review date on record",IF({L("Next due")}<TODAY(),"Overdue",'
                f'IF({L("Next due")}<TODAY()+30,"Due within 30 days","In date")))'))
            st.font, st.border, st.alignment = F(bold=True), BORDER, WRAP
            if d["file"]:
                for i in range(1, len(COLS)):
                    ws.cell(r, i).fill = fill(TINT[d["file"]])
            # No fixed height: Excel, Sheets and LibreOffice fit an unmerged row to its wrapped text.
            doc_rows.append(r)
            r += 1

    end = r - 1
    st = get_column_letter(COL["Status"])
    due = get_column_letter(COL["Next due"])
    red, amber, green, grey = fill("F4C7C3"), fill("FCE8B2"), fill("D7EBD5"), fill("E8E4E0")
    rng = f"{st}{HEADER_ROW + 1}:{st}{end}"
    for value, f_ in (("Overdue", red), ("Due within 30 days", amber), ("No review date on record", amber),
                      ("In date", green), ("Superseded", grey), ("Working copy", grey)):
        ws.conditional_formatting.add(rng, FormulaRule(formula=[f'{st}{HEADER_ROW + 1}="{value}"'], fill=f_))
    rng = f"{due}{HEADER_ROW + 1}:{due}{end}"
    ws.conditional_formatting.add(rng, FormulaRule(
        formula=[f"AND(ISNUMBER({due}{HEADER_ROW + 1}),{due}{HEADER_ROW + 1}<TODAY())"], fill=red))
    ws.conditional_formatting.add(rng, FormulaRule(
        formula=[f"AND(ISNUMBER({due}{HEADER_ROW + 1}),{due}{HEADER_ROW + 1}<TODAY()+30)"], fill=amber))
    ws.auto_filter.ref = f"A{HEADER_ROW}:{last}{end}"
    ws.freeze_panes = f"{get_column_letter(COL['type'])}{HEADER_ROW + 1}"
    ws.sheet_properties.outlinePr.summaryBelow = False
    # Printed: A3 landscape, one page wide, the checklist columns only (No. to Status).
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A3
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.print_title_rows = f"{HEADER_ROW}:{HEADER_ROW}"
    ws.print_area = f"A1:{st}{end}"
    return ws, end


SEND_COLS = [("No.", 6), ("Form (opens the file)", 46), ("They need to", 30), ("When it is needed", 44),
             ("Who fills it in", 32), ("Who signs", 32), ("Practice Standard", 30)]
NOT_SENT = ("Superseded", "Working copy")


def _send_sheet(wb, at, name, title, intro, tint, tab_colour, groups):
    """One tab: a title, then each group's forms under a band, every title linked."""
    ws = wb.create_sheet(name, at)
    ws.sheet_properties.tabColor = tab_colour
    for i, (_, w) in enumerate(SEND_COLS, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    last = get_column_letter(len(SEND_COLS))
    ws["A1"] = title
    ws["A1"].font = F(bold=True, size=14, color=FOREST)
    ws["A2"] = intro
    ws["A2"].font = F(color=MUTED)
    for i, (head, _) in enumerate(SEND_COLS, 1):
        c = ws.cell(4, i, head)
        c.font = F(bold=True, color="FFFFFF")
        c.fill = fill(FOREST)
        c.alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[4].height = 22
    r = 5
    for group, rows in groups:
        ws.merge_cells(f"A{r}:{last}{r}")
        c = ws.cell(r, 1, group)
        c.font = F(bold=True, size=11, color=FOREST)
        c.border = Border(top=Side(style="medium", color=FOREST))
        ws.row_dimensions[r].height = 24
        c.alignment = Alignment(vertical="bottom")
        r += 1
        for d, action in rows:
            values = [d["no"], d["title"], action, d["when"], d["completed_by"], d["signed_by"],
                      f"{d['module']}: {d['standard']}"]
            for i, v in enumerate(values, 1):
                c = ws.cell(r, i, v)
                c.font = F(bold=(i == 3))
                c.alignment = WRAP
                c.border = BORDER
                if tint:
                    c.fill = fill(tint)
            t = ws.cell(r, 2)
            if d["url"]:
                t.hyperlink = d["url"]
                t.font = F(color=FOREST, underline="single", bold=True)
            r += 1
    ws.freeze_panes = "C5"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.print_title_rows = "4:4"
    return r - 5


def write_send_tabs(wb, docs, at):
    """The 'Send to' tabs, then 'Harmony fills in' for every form on none of them."""
    by_title = {}
    for d in docs:
        if d["fixed_status"] not in NOT_SENT:
            by_title.setdefault(d["title"], []).append(d)
    sent = set()
    for i, tab in enumerate(SEND_TABS):
        groups = []
        for group, entries in tab["groups"]:
            rows = []
            for title, action in entries:
                if title not in by_title:
                    raise SystemExit(f"{tab['tab']}: no document called {title!r}")
                rows += [(d, action) for d in by_title[title]]
                sent.add(title)
            groups.append((group, rows))
        tint = TINT.get(tab["colour"]) if tab["colour"] != "Organisation" else None
        n = _send_sheet(wb, at + i, tab["tab"], tab["title"], tab["intro"], tint,
                        {"Participant": ROSE, "Worker": SAGE}.get(tab["colour"], FOREST), groups)
        print(f"  {tab['tab']}: {sum(len(g[1]) for g in groups)} rows")
    kept = {"Participant": [], "Worker": [], "Organisation": []}
    for d in docs:
        if d["file"] and d["title"] not in sent and d["fixed_status"] not in NOT_SENT:
            kept[d["file"]].append((d, "Harmony fills it in"))
    groups = [("Kept in each participant's file", kept["Participant"]),
              ("Kept in each worker's file", kept["Worker"]),
              ("The organisation's own records", kept["Organisation"])]
    _send_sheet(wb, at + len(SEND_TABS), "Harmony fills in", "Harmony fills these in itself (not sent out)",
                "Registers, records and plans the Principal, Director or Key Management Personnel complete and "
                "keep. Nobody outside Harmony needs to be sent these.", TINT["Organisation"], CREAM, groups)
    print(f"  Harmony fills in: {sum(len(g[1]) for g in groups)} rows")


def fix_start_here(wb, end):
    ws = wb["Start here"]
    st = get_column_letter(COL["Status"])
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, str) and "'All documents'!N2:N164" in c.value:
                c.value = c.value.replace("'All documents'!N2:N164", f"Checklist!{st}{HEADER_ROW + 1}:{st}{end}")
            elif c.value == "All documents":
                c.value = "Checklist"
                ws.cell(c.row, c.column + 1).value = (
                    "Laid out like the Avaana checklist: module, Practice Standard, its outcome and indicators, "
                    "then every document for it. Highlighted rows are forms someone must fill in: pink for each "
                    "participant's file, green for each worker's file, cream for the organisation's records. "
                    "Click a title to open the file. Filter by 'Filled in for' or 'Status'. The 'Send to' tabs "
                    "list what goes out to participants, workers and practitioners, and what each has to do; "
                    "'Harmony fills in' lists the forms that never leave the office.")
            elif isinstance(c.value, str) and c.value.startswith("After a review"):
                ws.cell(c.row, c.column + 1).value = (
                    "Type into the green-grey columns: the date in 'Last reviewed', and the version, approver and "
                    "approval date. Next due and Status update by themselves.")
            elif isinstance(c.value, str) and c.value.startswith("Built 16 October 2026"):
                c.value = c.value.replace("Built 16 October 2026",
                                          "Built 16 October 2026, set out by Practice Standard 19 October 2026,")
    ws.title = "Start here"


def main(base, out):
    wb, docs = read_base(base)
    enrich(docs)
    order = groups(docs)
    at = wb.sheetnames.index("All documents")
    del wb["All documents"]
    _, end = write_checklist(wb, order, at)
    write_send_tabs(wb, docs, at + 1)
    fix_start_here(wb, end)
    wb.active = 0
    wb.save(out)
    counts = {}
    for d in docs:
        counts[d["file"] or "Not filled in"] = counts.get(d["file"] or "Not filled in", 0) + 1
    print(f"{len(docs)} documents in {len(order)} groups -> {out}")
    print(counts)


if __name__ == "__main__":
    if len(sys.argv) != 3:
        raise SystemExit(__doc__)
    main(*sys.argv[1:])
