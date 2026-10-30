# Harmony forms

One Apps Script file, `Code.gs`, that turns the checklist workbook into the place Harmony sends
forms from and sees them come back. It adds a **Harmony forms** menu to the sheet.

- **People** tab: everyone forms go to (staff, participants, practitioners), with their email. A
  participant who does not sign for themselves has a guardian or representative and their email.
- **Send forms to someone**: pick the person, tick their forms (the list comes from their role's
  'Send to' tab), and a test copy of every email comes to the sender first. Then send.
- Every email goes out from the account that ran Set up, whoever in Harmony pressed the button: Send
  puts it on a queue that a one-minute timer empties. So every reply lands in the one inbox the hourly
  check reads. If the Send panel was closed, select the rows on 'Sent and returned' and choose Send
  the selected rows.
- A form whose signing template has its fields placed goes by **Google eSignature**: the script
  makes the person's copy and you press Request signature in it. Every other form goes by **email**
  as a PDF and a Word file, for them to sign and reply with.
- Every hour it files each signed copy that came back in the person's folder, dates 'Returned on' in
  'Sent and returned', and emails the people named in Settings.
- Every day it drafts a reminder (in Gmail drafts, never sent by the script) for any form out longer
  than 14 days. Every Monday it emails what is overdue or due soon on the Everything tab.
- **Bring the checklist up to date** makes an older copy of the checklist current from its own
  tabs: it adds the Everything tab and corrects three wordings. Set up runs it too. It sends nothing,
  and running it twice changes nothing the second time.
- **Mark reviewed today** stamps today's date as 'Last reviewed' for the document on the selected
  row; Next due and Status follow.

## Install

1. Put the checklist Excel in Drive, open it, and choose File > Save as Google Sheets.
2. In the Google Sheet: Extensions > Apps Script. Replace everything in `Code.gs` with this
   folder's `Code.gs`. Save.
3. Reload the sheet. Choose Harmony forms > Set up, and allow what Google asks for (Drive, Gmail,
   this sheet, and connecting to Drive's own API to turn Word forms into Google Docs).
4. Type people into the People tab, select their rows, and choose Harmony forms > Add the selected
   people.
5. Choose Harmony forms > Make signing templates. For each form that should go by eSignature, open
   its template, place the Signature, Name and Date fields (Tools > eSignature), and tick column D.

Emails go out from, and replies are read in, the Google account that runs Set up, and its timers run
as that account. Harmony's own account should run it, so the mail comes from Harmony and the signed
forms live in Harmony's Drive. If another account runs Set up later, it takes over: the earlier
account's timers stop themselves at their next run. Set up shares its two folders with everyone who
can edit the sheet; run it again after adding someone. Each person is asked once, the first time they
use the menu, to allow the script.

## Tests

```bash
node apps-script/forms/test/run.js
```

`test/fake.js` stands in for Sheets, Drive, Gmail and the timers, with the clock set just after
midnight in Brisbane so a time zone slip shows as a wrong day.

Not checked by the tests, because they need Google itself: Drive's conversion of a Word form to a
Google Doc, eSignature fields surviving a copy of the template, and the name Google gives the signed
PDF it files beside the copy. If that name does not contain the copy's name, a lone waiting form
still takes it, and with several waiting the script asks a person to match it.
