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
document states them. `hc-lockup.png` is the letterhead lockup, cropped unchanged
from the band the branded set is printed on.

Run with `/usr/bin/python3` (it has openpyxl). Re-running `make_workbook.py`
starts again from the 16 Oct file and drops anything typed since, so after the
first run only `export_web.py` is used.

`standards.json` is the wording of each outcome and indicator, taken from the
Rise platform's seed of the Quality Indicators Guidelines (compilation 3, in
force 1 July 2026). Placement rules (which document meets which standard) are
at the top of `make_workbook.py`. Who each form goes out to, and what they do
with it, is in `send_lists.py`; those lists become the Excel's 'Send to' tabs.
