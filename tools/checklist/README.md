# Documents and forms by Practice Standard

Harmony's compliance checklist in the shape of Avaana's 2023 NDIS application
checklist: module, Practice Standard, its outcome and indicators, then every
audited document for it. Forms someone must fill in are highlighted by whose
file they belong in (each participant, each worker, the organisation).

Two outputs from one workbook:

- **The Excel** (Harmony's working copy, kept in Drive, never in this repo). It
  holds the audit notes, so it stays private.
- **`site/checklist/`**, the same list on the Harmony Care site, filterable, with
  status worked out against today's date. This repo is public, so the export
  carries only what the page shows.

```bash
# once: lay the 16 Oct 2026 compliance calendar out by Practice Standard
python3 make_workbook.py "<16 Oct calendar>.xlsx" "<checklist>.xlsx"

# whenever Harmony updates the workbook: refresh the web page's data
python3 export_web.py "<checklist>.xlsx"

# the all-in printed version on Harmony's letterhead (uses harmony-care-docs/hcpdf.py)
python3 make_pdf.py "<checklist>.xlsx" "<harmony-care-docs folder>" "<out>.pdf"

# re-read each document's own version block (after downloading both audited folders with rclone)
python3 -I read_versions.py "<download folder>" brit.json m1sil.json
```

`versions.json` is that read: version, approver and approval date as each
document states them. `site/img/hc-lockup.png` is the letterhead lockup, cropped unchanged
from the band the branded set is printed on; the Excel and the team pages share it.

`everything.py` writes the workbook's 'Everything' tab and the PDF's Part 1: every document, staff
file item and year planner job under its Practice Standard. It reads dates from the other tabs, so
they are typed once. `make_pdf.py` needs Georgia, Harmony's document face, which ships with macOS; build
the PDF on a Mac.

Run with `/usr/bin/python3` (it has openpyxl). Re-running `make_workbook.py`
starts again from the 16 Oct file and drops anything typed since, so after the
first run only `export_web.py` is used.

`standards.json` is the wording of each outcome and indicator, taken from the
Rise platform's seed of the Quality Indicators Guidelines (compilation 3, in
force 1 July 2026). Placement rules (which document meets which standard) are
at the top of `make_workbook.py`. Who each form goes out to, and what they do
with it, is in `send_lists.py`; those lists become the Excel's 'Send to' tabs.

## Porting to Rise, or handing over to Harmony

Everything here is built to move. The website pages hold no data of their own:
they read `site/checklist/checklist.json` and `hub.json`, which `export_web.py`
writes from the workbook. The model, as Rise would hold it:

| Thing | Fields | Where it comes from now |
|---|---|---|
| Standard | Guidelines section, heading, outcome, indicators, Rules schedule and clause | `standards.json` (official text, compilation 3) |
| Document | number, title, type, Drive link, whose file (participant / worker / organisation), when needed, how often, completed by, signed by, version, approved by, approval date, next review, status | the workbook's Checklist tab; `versions.json` for versions |
| Placement | document to standard, plus "also meets" | `make_workbook.py` (placement rules, `SEE_ALSO`) |
| Destination | who a form goes to, the group, what they must do | `send_lists.py`, resolved by `destinations()` |
| Send | date sent, person, who they are, email, form, what to do, returned on, approved by, approved on | the workbook's "Sent and returned" tab (the site's "Copy rows for the tracker" gives the first six) |

If Rise is not ready in time, hand Harmony the workbook (the master), the PDF,
and this repo's `site/` folder; nothing in them depends on Rise or on Jesse's Drive
except the file links, which point at Harmony's own audited folders.
