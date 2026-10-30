"""The all-in printed version: everything by Practice Standard and the 'Send to' lists, on Harmony's letterhead.

Uses the page template of Harmony's document renderer (harmony-care-docs/hcpdf.py,
the one that printed the September 2026 branded set), so the cream ground,
header and footer bands and page numbering are exactly those documents'. Reads
the checklist workbook, the master, so the PDF says what the Excel says.

    python3 make_pdf.py CHECKLIST.xlsx RENDERER_DIR OUT.pdf
"""
import datetime as dt
import sys
from pathlib import Path

import openpyxl

import everything
import export_web
import make_workbook as mw
from send_lists import SEND_TABS


def main(xlsx, renderer, out):
    sys.path.insert(0, renderer)
    import hcpdf  # noqa: E402  (Harmony's renderer: fonts, page template, styles)
    from reportlab.lib.colors import HexColor
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import cm
    from reportlab.platypus import CondPageBreak, PageBreak, Paragraph, Spacer, Table, TableStyle

    docs = export_web.read(xlsx)
    modules = export_web.structure([dict(d) for d in docs])
    wb = openpyxl.load_workbook(xlsx)
    mw.check_never_named(wb)
    filed = everything.by_section(everything.staff_items(wb), everything.planner_jobs(wb, docs),
                                  [s["key"] for m in modules for s in m["standards"]])
    by_no = {d["no"]: d for d in docs}
    home = {n: s["heading"] for m in modules for s in m["standards"] for n in s["docs"]}
    today = dt.date.today()
    soon = (today + dt.timedelta(days=30)).isoformat()

    spec = {"type": "REGISTER", "label": "Checklist", "page": "A4L",
            "title": "Documents and forms by Practice Standard"}
    rail = mw.FOREST
    st = hcpdf.styles("#" + rail)
    avail = hcpdf.page_of(spec)["size"][0] - hcpdf.L - hcpdf.R - hcpdf.MEASURE
    MID, GREY, BODY = HexColor("#6A6A6A"), HexColor("#8C8C8C"), HexColor("#3A3A3A")
    small = ParagraphStyle("small", parent=st["cell"], fontSize=7.4, leading=10.2, textColor=MID)
    cell = ParagraphStyle("c", parent=st["cell"], fontSize=8, leading=11)
    cellb = ParagraphStyle("cb", parent=cell, fontName="Georgia-Bold")
    head = ParagraphStyle("hd", parent=st["cellh"], fontSize=6.6)
    module_st = ParagraphStyle("mod", fontName="Georgia-Bold", fontSize=15, leading=19,
                               textColor=HexColor("#" + rail), spaceBefore=6, spaceAfter=4)
    std_st = ParagraphStyle("std", parent=st["h2"], fontSize=11, spaceBefore=14, spaceAfter=1)
    cite_st = ParagraphStyle("cite", parent=small, fontName="Helvetica", fontSize=6.6, textColor=GREY,
                             spaceAfter=4)
    outcome_st = ParagraphStyle("out", parent=st["body"], fontName="Georgia-Italic", fontSize=9,
                                leading=12.6, spaceAfter=4)
    ind_st = ParagraphStyle("ind", parent=small, leftIndent=12, firstLineIndent=-12, spaceAfter=1.5)
    esc = hcpdf.esc

    def link(d, bold=False):
        title = esc(d["title"])
        if bold:
            title = f"<b>{title}</b>"
        if not d["url"]:
            return title
        return f'<link href="{esc(d["url"])}" color="#{rail}"><u>{title}</u></link>'

    def status(d):
        if d["fixed_status"] == "Missing":
            return "Missing", HexColor("#8A4B2A")
        if d["fixed_status"]:
            return d["fixed_status"], GREY
        due = d["next_due"]
        if not due:
            return "No review date on record", HexColor("#9A7B2F")
        if due < today.isoformat():
            return "Overdue", HexColor("#8A4B2A")
        if due < soon:
            return "Due within 30 days", HexColor("#9A7B2F")
        return "In date", HexColor("#" + rail)

    def pretty(iso):
        if not iso:
            return ""
        return dt.date.fromisoformat(iso).strftime("%-d %b %Y")

    def table(header, rows, widths, tints):
        data = [[Paragraph(h.upper(), head) for h in header]] + rows
        t = Table(data, colWidths=[avail * w for w in widths], repeatRows=1, hAlign="LEFT",
                  splitByRow=1)
        style = [
            ("BACKGROUND", (0, 0), (-1, 0), HexColor("#EEF1EA")),
            ("LINEBELOW", (0, 0), (-1, 0), 0.7, HexColor("#" + rail)),
            ("LINEBELOW", (0, 1), (-1, -1), 0.3, HexColor("#E6E2D8")),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING", (0, 0), (-1, -1), 5), ("RIGHTPADDING", (0, 0), (-1, -1), 5),
            ("TOPPADDING", (0, 0), (-1, -1), 4), ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ]
        for i, tint in enumerate(tints, 1):
            if tint:
                style.append(("BACKGROUND", (0, i), (-1, i), HexColor("#" + tint)))
        t.setStyle(TableStyle(style))
        return t

    story = [Paragraph(esc(spec["title"]), st["title"]),
             Paragraph("Everything Harmony Care holds, renews and does, set out under the NDIS Practice Standard "
                       "it meets: every document, every item each worker's file must hold, and every dated job, "
                       "with when each is needed, who it goes to, and whether it is in date. Status as at "
                       f"{today.strftime('%-d %B %Y')}.", st["sub"])]

    key = [[Paragraph("", cell), Paragraph("<b>Filled in for each participant</b>", cell)],
           [Paragraph("", cell), Paragraph("<b>Filled in for each worker</b>", cell)],
           [Paragraph("", cell), Paragraph("<b>Filled in by the organisation</b>", cell)],
           [Paragraph("", cell), Paragraph("No colour: policies and guides, read but not filled in", cell)]]
    kt = Table(key, colWidths=[0.9 * cm, 9 * cm], hAlign="LEFT")
    kt.setStyle(TableStyle([("BACKGROUND", (0, i), (0, i), HexColor("#" + c)) for i, c in
                            enumerate([mw.TINT["Participant"], mw.TINT["Worker"], mw.TINT["Organisation"],
                                       "FFFFFF"])] +
                           [("BOX", (0, 3), (0, 3), 0.3, HexColor("#D8D4CB")),
                            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                            ("TOPPADDING", (0, 0), (-1, -1), 3), ("BOTTOMPADDING", (0, 0), (-1, -1), 3)]))
    story += [kt, Spacer(1, 10)]
    parts = ["Part 1. Everything, by Practice Standard"] + \
            [f"Part {i + 2}. {t['title']}" for i, t in enumerate(SEND_TABS)] + \
            [f"Part {len(SEND_TABS) + 2}. Forms Harmony fills in itself (not sent out)"]
    story += [Paragraph("IN THIS DOCUMENT", st["h1"])] + [Paragraph(esc(p), st["li"]) for p in parts]
    story += [Spacer(1, 8), Paragraph(
        "Click any document title to open it. Outcome and indicator wording is quoted unchanged from the "
        "Federal Register of Legislation (legislation.gov.au): NDIS (Provider Registration and Practice "
        "Standards) Rules 2018, compilation 6 (F2026C00527), and NDIS (Quality Indicators for NDIS Practice "
        "Standards) Guidelines 2018, compilation 3 (F2026C00528), both in force from 1 July 2026.", small)]

    def doc_table(numbers):
        rows, tints = [], []
        for n in numbers:
            d = by_no[n]
            label, colour = status(d)
            rows.append([Paragraph(str(n), small), Paragraph(link(d, bool(d["file"])), cell),
                         Paragraph(esc(d["type"] + (f", version {d['version']}" if d["version"] else "")), small),
                         Paragraph(esc(mw.FILE_LABEL[d["file"]] if d["file"] else "Read, not filled in"),
                                   cellb if d["file"] else small),
                         Paragraph(esc(d["when"] or ""), small),
                         Paragraph(esc(pretty(d["next_due"])), small),
                         Paragraph(f'<font color="{colour.hexval()[2:] and "#" + colour.hexval()[2:]}">'
                                   f"<b>{esc(label)}</b></font>", small)])
            tints.append(mw.TINT[d["file"]] if d["file"] else None)
        return table(["No.", "Document", "Type", "Filled in for", "When it is needed", "Next due", "Status"],
                     rows, [.04, .25, .11, .1, .3, .08, .12], tints)

    def staff_table(items):
        rows = []
        for i in items:
            title = f'<link href="{esc(i["url"])}" color="#{rail}"><u>{esc(i["item"])}</u></link>' \
                if i["url"] else esc(i["item"])
            keep = esc(i["keep"]) + (f" <i>Harmony form: {esc(i['form'])}.</i>" if i["form"] else "")
            rows.append([Paragraph(title, cellb), Paragraph(esc(i["who"]), small),
                         Paragraph(esc(everything.WHEN_OF_GROUP.get(i["group"], i["group"])), small),
                         Paragraph(keep, small), Paragraph(esc(i["renew"] or ""), small)])
        return table(["Each worker's file", "Who needs it", "When it is needed", "What to keep", "Renew"],
                     rows, [.22, .12, .19, .33, .14], [mw.TINT["Worker"]] * len(rows))

    def job_table(jobs):
        rows = []
        for j in jobs:
            label, due = everything.job_status(j, today)
            colour = {"Overdue": HexColor("#8A4B2A"), "Due within 30 days": HexColor("#9A7B2F"),
                      "No record yet": HexColor("#9A7B2F")}.get(label, HexColor("#" + rail))
            when = j["due"] if j["due"] in everything.REPEATS else ""
            rows.append([Paragraph(f'<link href="{esc(j["url"])}" color="#{rail}"><u>{esc(j["what"])}</u></link>', cell),
                         Paragraph(esc(when or (due.strftime("%-d %b %Y") if due else "")), small),
                         Paragraph(f'<font color="#{colour.hexval()[2:]}"><b>{esc(label)}</b></font>', small),
                         Paragraph(esc(j["why"] or ""), small)])
        return table(["Organisation job", "Due", "Status", "Why"], rows, [.36, .1, .12, .42],
                     [mw.TINT["Organisation"]] * len(rows))

    # Part 1
    story += [PageBreak(), Paragraph("PART 1. EVERYTHING, BY PRACTICE STANDARD", st["h1"])]
    for m in modules:
        story += [CondPageBreak(5 * cm), Paragraph(esc(m["label"]), module_st)]
        for s in m["standards"]:
            story += [CondPageBreak(4.5 * cm), Paragraph(esc(s["heading"]), std_st)]
            if s["cite"]:
                story.append(Paragraph(esc(s["cite"]), cite_st))
            story.append(Paragraph(esc(("Outcome: " if s["cite"] else "") + s["outcome"]), outcome_st))
            for line in s["indicators"]:
                story.append(Paragraph(esc(line).replace("\n", "<br/>"), ind_st))
            for n in s["also"]:
                d = by_no[n]
                story.append(Paragraph(f"<i>Also meets this standard:</i> No. {n}, {link(d)} "
                                       f"(filed under {esc(home[n])})", small))
            if s["docs"]:
                story += [Spacer(1, 4), doc_table(s["docs"])]
            if filed[s["key"]]["staff"]:
                story += [Spacer(1, 4), staff_table(filed[s["key"]]["staff"])]
            if filed[s["key"]]["jobs"]:
                story += [Spacer(1, 4), job_table(filed[s["key"]]["jobs"])]

    # Parts 2 onwards: who each form goes to
    sent, kept = mw.destinations(docs)
    tabs = [(t["title"], t["intro"], mw.TINT.get(t["colour"]) if t["colour"] != "Organisation" else None, groups)
            for t, groups in sent]
    tabs.append(("Forms Harmony fills in itself (not sent out)",
                 "Registers, records and plans the Principal, Director or Key Management Personnel complete "
                 "and keep.", mw.TINT["Organisation"], kept))
    for i, (title, intro, tint, groups) in enumerate(tabs, 2):
        story += [PageBreak(), Paragraph(f"PART {i}. {esc(title).upper()}", st["h1"]),
                  Paragraph(esc(intro), st["covers"])]
        for g, rows in groups:
            if not rows:
                continue
            story += [CondPageBreak(3.5 * cm), Paragraph(esc(g), std_st), Spacer(1, 3)]
            body = [[Paragraph(str(d["no"]), small), Paragraph(link(d, True), cell),
                     Paragraph(f"<b>{esc(act)}</b>", cell), Paragraph(esc(d["when"] or ""), small),
                     Paragraph(esc(d["signed_by"] or ""), small)] for d, act in rows]
            story.append(table(["No.", "Form", "They need to", "When it is needed", "Who signs"], body,
                               [.04, .27, .19, .3, .2], [tint] * len(body)))

    import copy
    d1 = hcpdf.Doc(out, spec)
    d1.build(copy.deepcopy(story))
    d2 = hcpdf.Doc(out, spec)
    d2.total = d1.page
    d2.build(story)
    print(f"{d1.page} pages -> {out}")


if __name__ == "__main__":
    if len(sys.argv) != 4:
        raise SystemExit(__doc__)
    main(*sys.argv[1:])
