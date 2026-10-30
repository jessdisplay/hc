// node apps-script/forms/test/run.js
// Runs Code.gs against fake.js: a small Harmony workbook, then every path a form can take.
'use strict';
const assert = require('assert');
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { google, Blob } = require('./fake');

const CODE = fs.readFileSync(path.join(__dirname, '..', 'Code.gs'), 'utf8');
// 00:30 on Saturday 31 Oct in Brisbane, still Friday 30 Oct in UTC: a wrong time zone shows as a wrong day
const NOW = '2026-10-30T14:30:00Z';
const TODAY = '2026-10-31';

function world(now) {
  const g = google({ now: now || NOW });
  const ctx = vm.createContext(Object.assign({}, g.globals));
  vm.runInContext(CODE, ctx, { filename: 'Code.gs' });
  const { ss, drive } = g;
  const docx = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
  const audited = drive.folder('Documents', 'root');
  const forms = {
    decl: drive.file('Worker Declarations.docx', docx, audited.getId()),
    preEmp: drive.file('Pre Employment Collection Form.docx', docx, audited.getId()),
    consent: drive.file('Privacy Consent Form.docx', docx, audited.getId()),
    plan: drive.file('Mealtime Management Plan.docx', docx, audited.getId())
  };
  const sendTab = (name, band, rows) => {
    const s = ss.insertSheet(name);
    s.getRange(5, 1, 1, 7).setValues([['No.', 'Form (opens the file)', 'They need to', 'When it is needed',
      'Who fills it in', 'Who signs', 'Practice Standard']]);
    s.getRange(6, 1).setValue(band);
    rows.forEach(([no, f, need], i) => {
      s.getRange(7 + i, 1, 1, 3).setValues([[no, f.getName().replace('.docx', ''), need]]);
      s.cell(7 + i, 2).link = f.getUrl();
    });
  };
  sendTab('Send to workers', 'Before they start', [[53, forms.decl, 'Sign and return'],
                                                   [51, forms.preEmp, 'Fill in and sign']]);
  sendTab('Send to participants', 'Before supports start', [[4, forms.consent, 'Sign and return']]);
  sendTab('Send to practitioners', 'Qualified health practitioner', [[120, forms.plan, 'Write the plan, and sign']]);
  const tracker = ss.insertSheet('Sent and returned');
  tracker.getRange(5, 1, 1, 11).setValues([['Date sent', 'Sent to', 'Who they are', 'Email', 'Form',
    'They need to', 'Returned on', 'Approved by', 'Approved on', 'Status', 'Notes']]);
  const staff = ss.insertSheet('Staff file (copy per worker)');
  staff.getRange(1, 1).setValue('Staff file checklist');
  staff.getRange(3, 1, 3, 1).setValues([['Worker name'], ['Role'], ['Start date']]);
  staff.getRange(22, 2).setValue('100 points of ID');
  const team = ss.insertSheet('Team at a glance');
  team.getRange(4, 1).setValue('Worker (tab name)');
  const checklist = ss.insertSheet('Checklist');
  const head = Array.from({ length: 17 }, (_, i) => 'c' + (i + 1));
  head[0] = 'No.'; head[3] = 'Document (opens the file)'; head[14] = 'Last reviewed (type the date)';
  checklist.getRange(5, 1, 1, 17).setValues([head]);
  checklist.getRange(11, 1).setValue(1);
  checklist.getRange(11, 4).setValue('Promoting and Protecting Rights Policy');
  const every = ss.insertSheet('Everything');
  every.getRange(5, 1, 1, 12).setValues([['No.', 'Module', 'Practice Standard', 'What (opens the file)', 'Kind',
    'For', 'When it is needed', 'Details', 'How often', 'Next due', 'Status', 'Type dates in']]);
  every.getRange(6, 1).setValue('CORE MODULE');
  every.getRange(8, 1, 1, 12).setValues([[1, 'Core Module', 'Person-centred supports',
    'Promoting and Protecting Rights Policy', 'Policy and procedure', 'Not filled in', '', '', '',
    '31 Oct 2025', 'Overdue', 'Checklist, No. 1']]);
  every.getRange(9, 1, 1, 12).setValues([['', 'Core Module', 'Human Resource Management', 'First aid (HLTAID011)',
    "Worker's file", 'All workers', '', '', 'Every 3 years', '', 'Add workers on Team at a glance', '']]);
  every.getRange(10, 1, 1, 12).setValues([['', 'Core Module', 'Risk Management', 'Renew WorkCover', 'Dated job',
    'The organisation', '', '', '', '30 Nov 2026', 'Due within 30 days', 'Year planner']]);
  ss.active = every;
  return Object.assign(g, { ctx, forms, tracker, team, checklist, every, audited });
}

function addPerson(w, row, values) {
  const people = w.ss.getSheetByName('People');
  people.getRange(row, 1, 1, values.length).setValues([values]);
}

const tests = [];
const test = (name, fn) => tests.push({ name, fn });
const row = (w, r) => w.tracker.getRange(r, 1, 1, 18).getValues()[0];
const day = (d) => (d instanceof Date ? d.toISOString() : String(d));
const asDay = (w, d) => w.ctx.day_(d);
const plain = (x) => JSON.parse(JSON.stringify(x));  // values made inside the sandbox
/** every timer for this handler fires once, each as the account that created it, as Google does */
function fire(w, handler) {
  const was = w.user.effective;
  for (const t of w.triggers.filter((x) => x.handler === handler)) {
    w.user.effective = t.owner;
    w.ctx[handler]({ triggerUid: t.id });
  }
  w.user.effective = was;
}
/** what pressing Send does, end to end: the timer sends the test, Send queues, the timer sends it */
function deliver(w, rows) {
  w.ctx.sendQueued();
  const res = w.ctx.sendNow(rows);
  w.ctx.sendQueued();
  return res;
}

/** set up, three people, ready to send */
function ready(now) {
  const w = world(now);
  w.ctx.setup();
  addPerson(w, 5, ["Ana O'Brien", 'ana@example.com', 'Worker', '', '', '2026-11-03', 'Yes']);
  addPerson(w, 6, ['Sam Lee', '', 'Participant', 'Kim Lee', 'kim@example.com', '', 'Yes']);
  addPerson(w, 7, ['Dr Rae', 'rae@clinic', 'Practitioner or partner', '', '', '', 'Yes']);
  w.ss.select(w.ss.getSheetByName('People'), 5, 3);
  w.alerts.length = 0;
  w.ctx.addSelectedPeople();
  return w;
}

test('set up makes the tabs, folders and three timers, and running it twice changes nothing', () => {
  const w = world();
  w.ctx.setup();
  for (const t of ['People', 'Settings', 'Signing templates', 'Forms log']) assert.ok(w.ss.getSheetByName(t), t);
  assert.strictEqual(w.ss.tz, 'Australia/Brisbane');
  assert.strictEqual(w.ctx.setting_('Test copies go to'), 'jesse@riseos.care');
  assert.ok(w.ctx.folderId_(w.ctx.setting_("Folder for each person's files")));
  assert.deepStrictEqual(w.triggers.map((t) => t.handler).sort(), ['checkReturns', 'dailyChase', 'mondayDigest', 'sendQueued']);
  assert.strictEqual(w.tracker.getRange(5, 12).getValue(), 'Copy sent (opens it)');
  const folders = [...w.drive.items.values()].filter((x) => x.kind === 'folder').length;
  w.ctx.setup();
  assert.strictEqual(w.triggers.length, 4, 'timers not doubled');
  assert.strictEqual([...w.drive.items.values()].filter((x) => x.kind === 'folder').length, folders, 'no new folders');
  assert.strictEqual(w.ss.getSheetByName('Settings').getLastRow(), 10, 'settings not doubled');
  assert.strictEqual(w.ctx.setting_('Emails go out from'), 'jesse@riseos.care');
});

test("set up stops if 'Sent and returned' is not the workbook's", () => {
  const w = world();
  w.tracker.getRange(5, 5).setValue('Document');
  assert.throws(() => w.ctx.setup(), /column 5 should be 'Form'/);
});

test('adding people: a folder each, a worker gets a staff tab and a Team line, a bad address is reported', () => {
  const w = ready();
  assert.ok(w.ss.getSheetByName('Ana O’Brien'), 'tab named without the straight apostrophe');
  assert.strictEqual(w.ss.getSheetByName('Ana O’Brien').getRange(3, 2).getValue(), "Ana O'Brien");
  assert.strictEqual(w.team.getRange(5, 1).getValue(), 'Ana O’Brien');
  const people = w.ss.getSheetByName('People');
  for (const r of [5, 6, 7]) assert.ok(w.ctx.folderId_(people.getRange(r, 8).getValue()), 'folder for row ' + r);
  assert.match(w.alerts[0], /Ready: Ana O'Brien, Sam Lee, Dr Rae/);
  assert.match(w.alerts[0], /'rae@clinic' for Dr Rae is not an email address/);
  // adding again makes nothing twice
  const before = w.drive.items.size;
  w.ctx.addSelectedPeople();
  assert.strictEqual(w.drive.items.size, before);
  assert.strictEqual(w.team.getRange(6, 1).getValue(), '');
});

test('a full Team at a glance stops the menu item but never a send', () => {
  const w = ready();
  for (let r = 6; r <= 19; r++) w.team.getRange(r, 1).setValue('Worker ' + r);
  addPerson(w, 8, ['Bo Ng', 'bo@example.com', 'Worker', '', '', '', 'Yes']);
  w.ss.select(w.ss.getSheetByName('People'), 8, 1);
  w.alerts.length = 0;
  w.ctx.addSelectedPeople();
  assert.match(w.alerts[0], /Team at a glance is full/);
  const items = w.ctx.prepare(8, [w.forms.decl.getId()]);
  assert.ok(!items[0].error, items[0].error);
});

test("the sidebar lists a person's forms by their role's tab, all by email until a template is ready", () => {
  const w = ready();
  const d = w.ctx.formsFor(5);
  assert.deepStrictEqual(plain(d.to), ['ana@example.com']);
  assert.strictEqual(d.groups[0].name, 'Before they start');
  assert.deepStrictEqual(plain(d.groups[0].forms.map((f) => [f.title, f.how])),
    [['Worker Declarations', 'Email'], ['Pre Employment Collection Form', 'Email']]);
  const sam = w.ctx.formsFor(6);
  assert.deepStrictEqual(plain(sam.to), ['kim@example.com'], 'a participant without email: their guardian signs');
  assert.deepStrictEqual(plain(w.ctx.formsFor(7).problems), ["'rae@clinic' for Dr Rae is not an email address."]);
});

test('prepare copies the form into their folder, writes an undated row, and only the test copy goes', () => {
  const w = ready();
  const items = w.ctx.prepare(5, [w.forms.decl.getId()]);
  assert.strictEqual(items.length, 1);
  const it = items[0];
  assert.ok(!it.error, it.error);
  assert.strictEqual(it.how, 'Email');
  assert.strictEqual(it.tested, 'jesse@riseos.care');
  const v = row(w, it.trackerRow);
  assert.strictEqual(it.trackerRow, 6);
  assert.strictEqual(v[0], '', 'not dated until it really goes');
  assert.deepStrictEqual(plain(v.slice(1, 6)), ["Ana O'Brien", 'Worker', 'ana@example.com', 'Worker Declarations', 'Sign and return']);
  assert.strictEqual(v[12], 'Email');
  const copy = w.drive.get(w.ctx.fileId_(v[11]), 'file');
  assert.strictEqual(copy.mime, 'application/vnd.google-apps.document');
  assert.strictEqual(copy.name, "Worker Declarations - Ana O'Brien (31 Oct 2026)", 'Brisbane date, not UTC');
  assert.ok(w.ctx.folderId_(w.ss.getSheetByName('People').getRange(5, 8).getValue()), 'their folder link');
  assert.strictEqual(w.gmail.sent.length, 0, 'nothing goes until the timer runs');
  assert.deepStrictEqual(plain(v.slice(16, 18)), ['Test', 'jesse@riseos.care']);
  w.ctx.sendQueued();
  assert.strictEqual(w.gmail.sent.length, 1, 'only the test copy');
  assert.strictEqual(row(w, 6)[16], 'Test sent');
  const test = w.gmail.sent[0];
  assert.strictEqual(test.to, 'jesse@riseos.care');
  assert.match(test.subject, /^\[Test copy for Ana O'Brien\] Harmony Care: Worker Declarations$/);
  assert.match(test.body, /^Hi Ana,/);
  assert.match(test.body, /What to do: Sign and return\./);
  assert.deepStrictEqual(plain(test.opts.attachments.map((a) => a.name)), ['Worker Declarations.pdf', 'Worker Declarations.docx']);
  assert.ok(!/[\u2013\u2014]/.test(test.body), 'no dashes in what people read');
});

test('send goes once to the person, dates the row and keeps the thread; pressing again sends nothing', () => {
  const w = ready();
  const [it] = w.ctx.prepare(5, [w.forms.decl.getId()]);
  const res = deliver(w, [it.trackerRow]);
  assert.deepStrictEqual(plain(res), [{ row: 6, queued: true }]);
  const real = w.gmail.sent[1];
  assert.strictEqual(real.to, 'ana@example.com');
  assert.strictEqual(real.subject, 'Harmony Care: Worker Declarations');
  const v = row(w, 6);
  assert.strictEqual(asDay(w, v[0]), TODAY);
  assert.strictEqual(v[13], real.thread);
  assert.deepStrictEqual(plain(w.gmail.threads.get(real.thread).labels), ['Harmony forms']);
  assert.strictEqual(v[16], 'Sent');
  const again = deliver(w, [6]);
  assert.strictEqual(again[0].note, 'Already sent.');
  assert.strictEqual(w.gmail.sent.length, 2);
});

test('a participant with a guardian: the guardian gets it, greeted by name', () => {
  const w = ready();
  const [it] = w.ctx.prepare(6, [w.forms.consent.getId()]);
  deliver(w, [it.trackerRow]);
  const real = w.gmail.sent[w.gmail.sent.length - 1];
  assert.strictEqual(real.to, 'kim@example.com');
  assert.match(real.body, /^Hi Kim,/);
});

test('a bad address never gets a form', () => {
  const w = ready();
  assert.throws(() => w.ctx.prepare(7, [w.forms.plan.getId()]), /not an email address/);
  assert.strictEqual(w.gmail.sent.length, 0);
  assert.strictEqual(row(w, 6)[4], '', 'no row written');
});

test('the signed copy coming back by email is filed, dated and announced once', () => {
  const w = ready();
  const [it] = w.ctx.prepare(5, [w.forms.decl.getId()]);
  deliver(w, [it.trackerRow]);
  const thread = w.gmail.threads.get(row(w, 6)[13]);
  thread.add('Ana <ana@example.com>', 'jesse@riseos.care', 'Re', 'Done',
             [new Blob('signed', 'application/pdf', 'scan.pdf')]);
  w.ctx.checkReturns();
  const v = row(w, 6);
  assert.strictEqual(asDay(w, v[6]), TODAY);
  const saved = w.drive.get(w.ctx.fileId_(v[14]), 'file');
  assert.strictEqual(saved.name, "Worker Declarations - Ana O'Brien - signed.pdf");
  assert.strictEqual(saved.parent, w.ctx.folderId_(w.ss.getSheetByName('People').getRange(5, 8).getValue()));
  assert.strictEqual(w.mail.length, 1);
  assert.strictEqual(w.mail[0].subject, "Signed and back: Worker Declarations from Ana O'Brien");
  assert.strictEqual(w.mail[0].to, 'jesse@riseos.care');
  w.ctx.checkReturns();
  assert.strictEqual(w.mail.length, 1, 'not announced twice');
});

test('a reply without the form is noted and announced once, and the row stays open', () => {
  const w = ready();
  const [it] = w.ctx.prepare(6, [w.forms.consent.getId()]);
  deliver(w, [it.trackerRow]);
  const thread = w.gmail.threads.get(row(w, it.trackerRow)[13]);
  thread.add('kim@example.com', 'jesse@riseos.care', 'Re', 'Will do it Monday', []);
  w.ctx.checkReturns();
  w.ctx.checkReturns();
  const v = row(w, it.trackerRow);
  assert.strictEqual(v[6], '', 'still waiting');
  assert.match(v[10], /Replied without the signed form attached \(2026-10-31\)/);
  assert.strictEqual(w.mail.length, 1);
  assert.match(w.mail[0].subject, /^Reply without the form/);
});

test('our own sent email in the thread is never taken as the reply', () => {
  const w = ready();
  const [it] = w.ctx.prepare(5, [w.forms.decl.getId()]);
  deliver(w, [it.trackerRow]);
  w.ctx.checkReturns();
  assert.strictEqual(row(w, it.trackerRow)[6], '');
  assert.strictEqual(w.mail.length, 0);
});

test('templates: one Google Doc per form; a ticked one sends by eSignature and its signed PDF is matched', () => {
  const w = ready();
  w.ctx.makeTemplates();
  const t = w.ss.getSheetByName('Signing templates');
  assert.strictEqual(t.getLastRow(), 4 + 4, 'four distinct forms');
  w.alerts.length = 0;
  w.ctx.makeTemplates();
  assert.strictEqual(t.getLastRow(), 8, 'made once');
  const declRow = [5, 6, 7, 8].find((r) => t.getRange(r, 1).getValue() === 'Worker Declarations');
  assert.strictEqual(t.getRange(declRow, 4).getValue(), false);
  t.getRange(declRow, 4).setValue(true);
  assert.strictEqual(w.ctx.formsFor(5).groups[0].forms[0].how, 'Google eSignature');
  const [it] = w.ctx.prepare(5, [w.forms.decl.getId()]);
  assert.strictEqual(it.how, 'Google eSignature');
  assert.strictEqual(w.gmail.sent.length, 0, 'eSignature sends nothing by email');
  const copy = w.drive.get(w.ctx.fileId_(it.copyUrl), 'file');
  assert.strictEqual(copy.from, w.ctx.fileId_(t.getRange(declRow, 3).getValue()), 'copied from the template');
  w.ctx.checkReturns();
  assert.strictEqual(row(w, it.trackerRow)[6], '', 'not requested yet, nothing to find');
  w.ctx.markRequested(it.trackerRow);
  assert.strictEqual(asDay(w, row(w, it.trackerRow)[0]), TODAY);
  // Google files the completed PDF beside the copy
  w.drive.file(copy.name + ' - signed.pdf', 'application/pdf', copy.parent);
  w.ctx.checkReturns();
  assert.strictEqual(asDay(w, row(w, it.trackerRow)[6]), TODAY);
  assert.match(w.mail[0].subject, /^Signed and back: Worker Declarations/);
});

test('eSignature: two forms waiting and a PDF named for neither is handed to a person, never guessed', () => {
  const w = ready();
  w.ctx.makeTemplates();
  const t = w.ss.getSheetByName('Signing templates');
  for (let r = 5; r <= 8; r++) t.getRange(r, 4).setValue(true);
  const items = w.ctx.prepare(5, [w.forms.decl.getId(), w.forms.preEmp.getId()]);
  items.forEach((it) => w.ctx.markRequested(it.trackerRow));
  const folder = w.drive.get(w.ctx.fileId_(items[0].copyUrl), 'file').parent;
  w.drive.file('Completed document.pdf', 'application/pdf', folder);
  w.ctx.checkReturns();
  items.forEach((it) => assert.strictEqual(row(w, it.trackerRow)[6], ''));
  assert.match(w.mail[0].subject, /^A signed file to match/);
  w.ctx.checkReturns();
  assert.strictEqual(w.mail.length, 1, 'asked once');
});

test("eSignature: one form waiting takes the folder's new PDF; an emailed signed copy is never taken", () => {
  const w = ready();
  w.ctx.makeTemplates();
  const t = w.ss.getSheetByName('Signing templates');
  const declRow = [5, 6, 7, 8].find((r) => t.getRange(r, 1).getValue() === 'Worker Declarations');
  t.getRange(declRow, 4).setValue(true);
  const [byEmail] = w.ctx.prepare(5, [w.forms.preEmp.getId()]);
  deliver(w, [byEmail.trackerRow]);
  const [bySign] = w.ctx.prepare(5, [w.forms.decl.getId()]);
  w.ctx.markRequested(bySign.trackerRow);
  w.gmail.threads.get(row(w, byEmail.trackerRow)[13])
    .add('ana@example.com', 'me', 'Re', 'Here', [new Blob('x', 'application/pdf', 'pre.pdf')]);
  w.ctx.checkReturns();
  assert.strictEqual(asDay(w, row(w, byEmail.trackerRow)[6]), TODAY);
  assert.strictEqual(row(w, bySign.trackerRow)[6], '', 'the emailed PDF is not the eSignature return');
  const folder = w.drive.get(w.ctx.fileId_(bySign.copyUrl), 'file').parent;
  w.drive.file('Signed.pdf', 'application/pdf', folder);
  w.ctx.checkReturns();
  assert.strictEqual(asDay(w, row(w, bySign.trackerRow)[6]), TODAY);
});

test('chasing: a form out 14 days gets a Gmail draft (never sent) and one list, then waits a week', () => {
  const w = ready();
  const [it] = w.ctx.prepare(5, [w.forms.decl.getId()]);
  deliver(w, [it.trackerRow]);
  w.tracker.getRange(it.trackerRow, 1).setValue('2026-10-16');
  w.ctx.dailyChase();
  assert.strictEqual(w.gmail.drafts.length, 1);
  assert.match(w.gmail.drafts[0].body, /we sent on 16 October/);
  assert.strictEqual(w.gmail.sent.length, 2, 'the reminder is a draft only');
  assert.match(w.mail[0].subject, /^1 form to chase/);
  w.ctx.dailyChase();
  assert.strictEqual(w.gmail.drafts.length, 1, 'not again within the week');
});

test('chasing leaves a form alone before its 14 days', () => {
  const w = ready();
  const [it] = w.ctx.prepare(5, [w.forms.decl.getId()]);
  deliver(w, [it.trackerRow]);
  w.tracker.getRange(it.trackerRow, 1).setValue('2026-10-18');
  w.ctx.dailyChase();
  assert.strictEqual(w.gmail.drafts.length, 0);
  assert.strictEqual(w.mail.length, 0);
});

test('mark reviewed today: from the Everything row to the Checklist cell, in Brisbane time', () => {
  const w = ready();
  w.ss.select(w.every, 8);
  w.ctx.markReviewedToday();
  assert.strictEqual(asDay(w, w.checklist.getRange(11, 15).getValue()), TODAY);
  assert.match(w.ss.toasts[0], /No\. 1 marked reviewed today/);
  w.ss.select(w.every, 9);
  w.alerts.length = 0;
  w.ctx.markReviewedToday();
  assert.match(w.alerts[0], /Select a document row/);
});

test('the Monday email counts what is overdue, due soon, out and waiting', () => {
  const w = ready();
  const [it] = w.ctx.prepare(5, [w.forms.decl.getId()]);
  deliver(w, [it.trackerRow]);
  w.ctx.mondayDigest();
  const m = w.mail[0];
  assert.strictEqual(m.subject, 'Harmony Care this week: 1 overdue or missing, 1 due soon');
  assert.match(m.body, /Promoting and Protecting Rights Policy \(Person-centred supports\), due 31 Oct 2025: Overdue/);
  assert.match(m.body, /Renew WorkCover/);
  assert.match(m.body, /Forms out and not back: 1\./);
});

test("Drive refusing a copy is reported on that form and leaves no row behind", () => {
  const w = ready();
  w.failFetch((url) => (url.includes(w.forms.preEmp.getId()) ? 403 : 200));
  const items = w.ctx.prepare(5, [w.forms.preEmp.getId(), w.forms.decl.getId()]);
  assert.match(items[0].error, /Google could not make a Google Doc .*403/);
  assert.ok(!items[1].error);
  assert.strictEqual(items[1].trackerRow, 6, 'the failed one took no row');
  assert.strictEqual(w.ss.getSheetByName('Forms log').getLastRow() > 3, true);
});

test('tab names: no characters Sheets refuses and no straight apostrophe', () => {
  const w = world();
  assert.strictEqual(w.ctx.sheetName_("Jo/Ann [O'Neil]*"), 'Jo Ann O’Neil');
});

test('nothing breaks at the turn of the day: a date written is the date read', () => {
  for (const at of ['2026-10-30T13:59:00Z', '2026-10-30T14:00:00Z', '2026-12-31T23:30:00Z']) {
    const w = ready(at);
    const today = w.ctx.day_(new w.clock());
    const cell = w.tracker.getRange(400, 1);
    w.ctx.date_(cell, new w.clock());
    assert.strictEqual(w.ctx.day_(cell.getValue()), today, at);
  }
});

test('Brittany sends: her test copy comes to her, the email goes from the Set up account, the reply is caught', () => {
  const w = ready();
  w.user.active = 'brittany@harmony.example';
  const d = w.ctx.formsFor(5);
  assert.strictEqual(d.testTo, 'brittany@harmony.example', 'her own address offered');
  assert.strictEqual(d.from, 'jesse@riseos.care');
  const [it] = w.ctx.prepare(5, [w.forms.decl.getId()], 'brittany@harmony.example');
  w.ctx.sendQueued();
  assert.strictEqual(w.gmail.sent[0].to, 'brittany@harmony.example');
  w.ctx.sendNow([it.trackerRow]);
  w.ctx.sendQueued();
  const real = w.gmail.sent[1];
  assert.strictEqual(real.to, 'ana@example.com');
  w.gmail.threads.get(real.thread).add('ana@example.com', 'jesse@riseos.care', 'Re', 'Signed',
                                       [new Blob('s', 'application/pdf', 'signed.pdf')]);
  w.ctx.checkReturns();
  assert.strictEqual(asDay(w, row(w, it.trackerRow)[6]), TODAY, 'caught although Brittany pressed Send');
  assert.strictEqual(w.ctx.testTo_(), 'brittany@harmony.example', 'remembered for her');
  w.user.active = 'jesse@riseos.care';
  assert.strictEqual(w.ctx.testTo_(), 'jesse@riseos.care', 'not for anyone else');
});

test('Send before the test copy has gone is refused, and nothing reaches the person', () => {
  const w = ready();
  const [it] = w.ctx.prepare(5, [w.forms.decl.getId()]);
  const res = w.ctx.sendNow([it.trackerRow]);
  assert.match(res[0].error, /test copy has not gone yet/);
  w.ctx.sendQueued();
  assert.deepStrictEqual(plain(w.gmail.sent.map((m) => m.to)), ['jesse@riseos.care'], 'only the test');
  assert.strictEqual(row(w, it.trackerRow)[16], 'Test sent');
});

test('a send that fails in the queue says why on its row and tells Settings, once', () => {
  const w = ready();
  const [it] = w.ctx.prepare(5, [w.forms.decl.getId()]);
  w.ctx.sendQueued();
  w.ss.getSheetByName('People').getRange(5, 2).setValue('ana at example');
  w.ctx.sendNow([it.trackerRow]);
  w.ctx.sendQueued();
  w.ctx.sendQueued();
  const v = row(w, it.trackerRow);
  assert.match(v[16], /^Not sent: 'ana at example' for Ana O'Brien is not an email address/);
  assert.strictEqual(v[0], '', 'not dated');
  assert.strictEqual(w.gmail.sent.length, 1, 'only the test went');
  assert.strictEqual(w.mail.filter((m) => /^Not sent/.test(m.subject)).length, 1);
});

test('a reply without the form tells Brittany what they wrote, since it landed in another inbox', () => {
  const w = ready();
  const [it] = w.ctx.prepare(5, [w.forms.decl.getId()]);
  deliver(w, [it.trackerRow]);
  w.gmail.threads.get(row(w, it.trackerRow)[13]).add('ana@example.com', 'me', 'Re', 'Which page do I sign?', []);
  w.ctx.checkReturns();
  assert.match(w.mail[0].body, /They wrote:\n\nWhich page do I sign\?/);
});

test('the menu item sends selected rows when the panel was closed', () => {
  const w = ready();
  const [it] = w.ctx.prepare(5, [w.forms.decl.getId()]);
  w.ctx.sendQueued();
  w.ss.select(w.tracker, it.trackerRow);
  w.alerts.length = 0;
  w.ctx.sendSelectedRows();
  assert.match(w.alerts[0], /^1 on the way/);
  w.ctx.sendQueued();
  assert.strictEqual(w.gmail.sent[1].to, 'ana@example.com');
});

test("Brittany runs Set up: emails go from her, and Jesse's leftover timers stop themselves", () => {
  const w = ready();  // Jesse ran Set up first
  w.user.active = w.user.effective = 'brittany@harmony.example';
  w.alerts.length = 0;
  w.ctx.setup();
  assert.match(w.alerts[0], /Emails now go out from brittany@harmony\.example\. jesse@riseos\.care no longer sends/);
  assert.strictEqual(w.ctx.setting_('Emails go out from'), 'brittany@harmony.example');
  assert.strictEqual(w.triggers.filter((t) => t.owner === 'jesse@riseos.care').length, 4, 'still there until they fire');
  const [it] = w.ctx.prepare(5, [w.forms.decl.getId()], 'brittany@harmony.example');
  fire(w, 'sendQueued');
  assert.strictEqual(w.triggers.filter((t) => t.owner === 'jesse@riseos.care').length, 0, "Jesse's timers removed");
  assert.strictEqual(w.triggers.filter((t) => t.owner === 'brittany@harmony.example').length, 4);
  assert.deepStrictEqual(plain(w.gmail.sent.map((m) => [m.to, m.from])),
    [['brittany@harmony.example', 'brittany@harmony.example']], 'the test went once, from Brittany');
  w.ctx.sendNow([it.trackerRow]);
  fire(w, 'sendQueued');
  assert.deepStrictEqual(plain(w.gmail.sent[1].from), 'brittany@harmony.example');
  assert.strictEqual(w.gmail.sent[1].to, 'ana@example.com');
  fire(w, 'checkReturns');
  fire(w, 'dailyChase');
  fire(w, 'mondayDigest');
  assert.ok(w.mail.every((m) => m.from === 'brittany@harmony.example'), 'nothing from Jesse after the move');
});

test('anyone who is not the sending account gets a plain answer from Check for returned forms', () => {
  const w = ready();
  w.user.active = w.user.effective = 'victoria@harmony.example';
  w.alerts.length = 0;
  w.ctx.checkReturns();
  assert.match(w.alerts[0], /checked from jesse@riseos\.care's account/);
  assert.strictEqual(w.triggers.length, 4, "a menu click removes nobody's timers");
});

test("Set up shares the two folders with everyone who can edit the sheet, once, and says so", () => {
  const w = world();
  w.ss.editors = ['brittany@harmony.example', 'victoria@harmony.example', 'jesse@riseos.care'];
  w.user.active = w.user.effective = 'brittany@harmony.example';
  w.ctx.setup();
  assert.match(w.alerts[0], /shared with: victoria@harmony\.example, jesse@riseos\.care\./);
  const folder = w.drive.get(w.ctx.folderId_(w.ctx.setting_("Folder for each person's files")), 'folder');
  assert.strictEqual(folder.owner, 'brittany@harmony.example', "in Brittany's Drive");
  assert.deepStrictEqual(plain(folder.editors), ['victoria@harmony.example', 'jesse@riseos.care']);
  w.alerts.length = 0;
  w.ctx.setup();
  assert.ok(!/shared with/.test(w.alerts[0]), 'nobody added twice');
  assert.deepStrictEqual(plain(folder.editors), ['victoria@harmony.example', 'jesse@riseos.care']);
});

let failed = 0;
for (const t of tests) {
  try {
    t.fn();
    console.log('ok   ' + t.name);
  } catch (err) {
    failed++;
    console.log('FAIL ' + t.name + '\n     ' + String(err.stack || err).split('\n').slice(0, 4).join('\n     '));
  }
}
console.log(`\n${tests.length - failed} of ${tests.length} passed`);
process.exit(failed ? 1 : 0);
