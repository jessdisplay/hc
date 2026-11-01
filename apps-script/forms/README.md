# Harmony forms

The engine behind Harmony's compliance tool. `Code.gs` lives on the checklist Google Sheet, which is the
store: nobody at Harmony needs to open it. They use **hc.riseos.care/tool** (the `tool/` folder of this
repo): Today, Checklist and Forms. The page holds no data; it signs in with a code emailed to the person
and reads and writes through `doPost`, the script's only way in.

- **Sign in:** only addresses on the Settings line 'Who can sign in to the tool' get a code. Six digits,
  10 minutes, 5 tries, 5 codes an hour. A session lasts 14 days; only its hash is stored, and taking
  someone off the line ends their session at once.
- **Forms:** the list for each person comes from their role's 'Send to' tab. Tick what goes out; a test
  copy comes to whoever pressed the button; then send. Forms handed over in person are recorded the same
  way. Each one shows where it stands: not sent, with them, chase up, back to check, done.
- **Read-only documents** ('Read', 'Read, then confirm...') go as one email of links, with no copy made,
  nothing to return and nothing to chase. 'Read and sign' comes back like a form.
- **Signed forms:** a form whose signing template has its fields placed goes by Google eSignature, and
  you press Request signature in the copy. Every other form goes by email as a PDF and a Word file.
  Every hour the script files each signed copy that came back in the person's folder and dates its row.
- **Email:** everything goes from the account that ran Set up, through a queue a one-minute timer empties,
  so every reply lands in the one inbox the hourly check reads. Every day it drafts a reminder in Gmail
  (never sent by the script) for a form out longer than 14 days. Every Monday it emails what is due.

## Install

1. In the Google Sheet: Extensions > Apps Script. Replace everything in `Code.gs` with this folder's
   `Code.gs`. Save.
2. Reload the sheet. Choose Harmony forms > Set up, and allow what Google asks for. Set up also sets the
   Sheet to Brisbane time.
3. In the Apps Script editor: Deploy > New deployment > Web app. Execute as: Me. Who has access: Anyone.
   Copy the address ending in `/exec`.
4. Put that address in `tool/config.js` and publish.
5. On the Settings tab, list who can sign in. For eSignature: Harmony forms > Make signing templates, place
   the fields in each template, and tick column D.

"Anyone" means anyone can reach the sign-in, not the data. Every call past the sign-in checks a session, and
the script serves no page of its own, so `doPost` is the only door. A change to `Code.gs` reaches the tool
only through Deploy > Manage deployments > Edit > New version.

The account that runs Set up sends every email and reads the replies. Harmony's own account should, so
the mail comes from Harmony and the signed forms live in Harmony's Drive. If another account runs Set up
later it takes over, and must deploy again. Set up shares its two folders with everyone who can edit the
sheet.

## Tests

```bash
node apps-script/forms/test/run.js
```

`test/fake.js` stands in for Sheets, Drive, Gmail and the timers, with the clock set just after
midnight in Brisbane so a time zone slip shows as a wrong day.

Not checked by the tests, because they need Google itself: `doPost` over the web and the page reaching it
from another address, Drive's conversion of a Word form to a
Google Doc, eSignature fields surviving a copy of the template, and the name Google gives the signed
PDF it files beside the copy. If that name does not contain the copy's name, a lone waiting form
still takes it, and with several waiting the script asks a person to match it.
