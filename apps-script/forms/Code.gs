/**
 * Harmony Care forms :: send forms to people to sign, and see them come back
 *
 * The engine behind Harmony's tool (hc.riseos.care/tool). The Sheet is the store: people use the tool,
 * never the tabs. doPost at the bottom of this file is the tool's only way in. Set up and the office
 * jobs stay on a 'Harmony forms' menu in the Sheet:
 *
 *   Set up                    first time only: People, Settings, Signing templates and Forms log tabs,
 *                             the two Drive folders, and the hourly, daily and Monday timers
 *   Add the selected people   a Drive folder each; a worker also gets their own staff file tab and a
 *                             line on Team at a glance
 *   Send the selected rows    on 'Sent and returned': send tested rows, if the Send panel was closed
 *   Make signing templates    a Google Doc of every form, for placing eSignature fields once
 *   Check for returned forms  what the hourly timer does, now
 *   Mark reviewed today       stamps today's date as 'Last reviewed' for the document on the selected row
 *   Email me what is due      what the Monday email says, now
 *
 * Three ways a form goes out. The sign page is the main one: one link in one email, they fill in and
 * sign in the browser, and it is filed the moment they finish. Or a form whose signing template has its
 * fields placed goes by Google eSignature: the script makes the person's copy, you press Request
 * signature in it (Google has no way for a script to press it), and the signed PDF Google files beside
 * the copy is picked up by the hourly check. Or it goes by email as a PDF, and the signed copy they reply
 * with is filed the same way. Every way 'Sent and returned' gets the row, the person's folder gets the
 * file, and the people in Settings get an email when it is back.
 *
 * Every email goes out from the account that ran Set up, never from the person pressing the button.
 * Pressing Send puts it on a queue that a one-minute timer (running as that account) empties. That
 * way every reply lands in the one inbox the hourly check reads, whoever in Harmony sent the form.
 *
 * Install: Extensions > Apps Script, replace Code.gs with this file, save, reload the sheet, then
 * Harmony forms > Set up. Google asks once for permission to use Drive, Gmail and this sheet.
 */

var TZ = 'Australia/Brisbane';
var TABS = {
  people: 'People', settings: 'Settings', templates: 'Signing templates', log: 'Forms log',
  tracker: 'Sent and returned', staff: 'Staff file (copy per worker)', team: 'Team at a glance',
  checklist: 'Checklist', everything: 'Everything'
};
var ROLES = ['Participant', 'Worker', 'Practitioner or partner'];
var SEND_TABS = {
  'Participant': 'Send to participants', 'Worker': 'Send to workers',
  'Practitioner or partner': 'Send to practitioners'
};
var LABEL = 'Harmony forms';  // every email thread the script starts carries this Gmail label
var ESIGN = 'Google eSignature', EMAIL = 'Email', BY_HAND = 'By hand', SIGN_PAGE = 'Sign page';
var MARK = 'Harmony forms: ';  // on a file's description: the hourly check has dealt with it

// People: one row per person, typed in by Harmony.
var PEOPLE_HEAD = 4;
var PEOPLE = ['Name', 'Email', 'Role', 'Signs for them (guardian or representative)', 'Their email',
              'Start date', 'Active', 'Folder', 'Notes'];
var P = index_(PEOPLE);

// Sent and returned: the workbook's own columns A to K, then the seven this script keeps.
var TRACKER_HEAD = 5, TRACKER_FIRST = 6;
var TRACKER = ['Date sent', 'Sent to', 'Who they are', 'Email', 'Form', 'They need to', 'Returned on',
               'Approved by', 'Approved on', 'Status', 'Notes',
               'Copy sent (opens it)', 'How it went out', 'Email thread', 'Signed copy', 'Last chased',
               'Send step', 'Test copy to'];
// Send step: 'Test' (queued) > 'Test sent' > 'Send' (queued) > 'Sent', or 'Not sent: why'.
var T = index_(TRACKER);

var TEMPLATE_HEAD = 4;
var TEMPLATES = ['Form', 'Goes to', 'Signing template (open it to place the fields)',
                 'Signature fields placed', 'Original'];
var TP = index_(TEMPLATES);

var SETTINGS = [
  ['Send emails as', 'Harmony Care', 'The name people see the emails come from.'],
  ['Test copies go to', '', 'Where a test copy goes if the sender gives no address in the Send panel.'],
  ['Emails go out from', '', 'Set up fills this: the account that ran it. Every email goes from here.'],
  ['Tell these people when a form comes back', '', 'Comma separated. They also get the Monday email.'],
  ['Who can sign in to the tool', '', 'Comma separated email addresses. Each gets a code by email to sign in.'],
  ["Folder for each person's files", '', 'Set up makes it. Each person gets a folder inside it.'],
  ['Signing templates folder', '', 'Set up makes it. The signing templates live here.'],
  ['Chase after (days)', 14, 'A form not back after this many days gets a reminder drafted in Gmail.']
];

var TOOL_URL = 'https://hc.riseos.care/tool/';  // where people work; every notice points here, never at a tab
var SIGN_URL = TOOL_URL + 'sign/#', LINK_DAYS = 90;
var RECORDS = {
  root: 'Harmony Care records', all: 'All completed documents', sent: 'Sent out', done: 'Completed', staff: 'Staff file',
  role: { 'Participant': 'Participants', 'Worker': 'Workers', 'Practitioner or partner': 'Practitioners and partners' }
};
var TIME_BUDGET_MS = 4.5 * 60 * 1000;  // Apps Script stops a run at 6 minutes
var TIMERS = ['sendQueued', 'checkReturns', 'dailyChase', 'mondayDigest'];


/* ------------------------------------------------------------------ menu and set up */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Harmony forms')
    .addItem('Send the selected rows', 'sendSelectedRows')
    .addItem('Add the selected people', 'addSelectedPeople')
    .addSeparator()
    .addItem('Check for returned forms', 'checkReturns')
    .addItem('Mark reviewed today', 'markReviewedToday')
    .addItem('Email me what is due', 'mondayDigest')
    .addSeparator()
    .addItem('Make signing templates', 'makeTemplates')
    .addItem('Bring the checklist up to date', 'upgradeWorkbook')
    .addItem('Set up', 'setup')
    .addToUi();
}

function setup() {
  var ss = SpreadsheetApp.getActive();
  PropertiesService.getScriptProperties().setProperty('SHEET_ID', ss.getId());
  ss.setSpreadsheetTimeZone(TZ);
  upgrade_(ss);
  var me = Session.getEffectiveUser().getEmail();

  var settings = ensureTab_(ss, TABS.settings, 'Settings',
    'What the forms menu uses. Change a value in column B; leave the labels in column A as they are.',
    3, ['Setting', 'Value', 'What it does']);
  var labels = settings.getRange(4, 1, Math.max(settings.getLastRow() - 3, 1), 1).getValues()
    .map(function (r) { return r[0]; });
  SETTINGS.forEach(function (row) {
    if (labels.indexOf(row[0]) === -1) settings.appendRow(row);
  });
  var before = String(setting_('Emails go out from') || '');
  putSetting_('Emails go out from', me);
  if (!setting_('Test copies go to')) putSetting_('Test copies go to', me);
  if (!setting_('Tell these people when a form comes back')) {
    putSetting_('Tell these people when a form comes back', me);
  }
  if (!setting_('Who can sign in to the tool')) putSetting_('Who can sign in to the tool', me);
  if (!folderId_(setting_("Folder for each person's files"))) {
    putSetting_("Folder for each person's files",
                DriveApp.createFolder(RECORDS.root).getUrl());
  }
  if (!folderId_(setting_('Signing templates folder'))) {
    putSetting_('Signing templates folder', DriveApp.createFolder('Harmony Care - signing templates').getUrl());
  }
  var root = recordsRoot_();
  root.setName(RECORDS.root);
  allCompleted_();
  ROLES.forEach(function (role) { childFolder_(root, RECORDS.role[role]); });
  var shared = shareFolders_(ss);
  var opening = learnKinds_();

  var people = ensureTab_(ss, TABS.people, 'People',
    'Everyone Harmony sends forms to: staff, participants, practitioners. Type a name, email and role, ' +
    'then select the row and choose Harmony forms > Add the selected people. A participant who does not ' +
    'sign for themselves: put the guardian or representative and their email in the next two columns.',
    PEOPLE_HEAD, PEOPLE);
  var max = Math.max(people.getMaxRows(), 200);
  people.getRange(PEOPLE_HEAD + 1, P['Role'] + 1, max - PEOPLE_HEAD, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(ROLES, true).build());
  people.getRange(PEOPLE_HEAD + 1, P['Active'] + 1, max - PEOPLE_HEAD, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['Yes', 'No'], true).build());

  ensureTab_(ss, TABS.templates, 'Signing templates',
    'A Google Doc of each form that goes out. Open one, choose Tools > eSignature, place a Signature field ' +
    '(and Name and Date) for each person who signs, then tick column D. From then on that form goes by ' +
    'Google eSignature. A form without a ticked template goes by email for them to sign and send back.',
    TEMPLATE_HEAD, TEMPLATES);
  ensureTab_(ss, TABS.log, 'Forms log', 'What the forms menu and its timers did, newest at the bottom.',
    3, ['When', 'What', 'Detail']);

  var tracker = tab_(TABS.tracker);
  var head = tracker.getRange(TRACKER_HEAD, 1, 1, TRACKER.length);
  var have = head.getValues()[0];
  for (var i = 0; i < 11; i++) {
    if (have[i] !== TRACKER[i]) throw new Error("'Sent and returned' has changed: column " + (i + 1) +
                                                " should be '" + TRACKER[i] + "'.");
  }
  head.setValues([TRACKER]);
  tracker.getRange(TRACKER_HEAD, 12, 1, TRACKER.length - 11)
    .setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#38553E').setWrap(true);

  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (TIMERS.indexOf(t.getHandlerFunction()) !== -1) {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('sendQueued').timeBased().everyMinutes(1).create();
  ScriptApp.newTrigger('checkReturns').timeBased().everyMinutes(15).create();
  ScriptApp.newTrigger('dailyChase').timeBased().everyDays(1).atHour(8).inTimezone(TZ).create();
  ScriptApp.newTrigger('mondayDigest').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(8)
    .inTimezone(TZ).create();

  var moved = before && before.toLowerCase() !== me.toLowerCase();
  log_('Set up', opening + ' files open in Google Docs or Sheets. Tabs, folders and timers ready. Emails go out from ' + me +
       (moved ? '; ' + before + "'s timers stop at their next run." : '.'));
  ui_('Set up is done. Emails now go out from ' + me + '.' +
      (moved ? ' ' + before + ' no longer sends anything.' : '') +
      (shared.length ? '\n\nThe two forms folders are now shared with: ' + shared.join(', ') + '.' : '') +
      '\n\nNext: deploy the web app, then add people in the tool: ' + TOOL_URL);
}


/* ------------------------------------------------------------------ people */

/** The menu item: a folder for every selected person; a worker also gets a staff file tab. */
function addSelectedPeople() {
  var sheet = SpreadsheetApp.getActiveSheet();
  if (sheet.getName() !== TABS.people) return ui_('Select the people on the People tab first.');
  var range = sheet.getActiveRange();
  var done = [], problems = [];
  for (var r = range.getRow(); r < range.getRow() + range.getNumRows(); r++) {
    if (r <= PEOPLE_HEAD) continue;
    try {
      var person = person_(r);
      if (!person.name) continue;
      setUpPerson_(person, true);
      done.push(person.name);
      var who = signers_(person);
      if (who.problems.length) problems.push('Row ' + r + ': ' + who.problems.join(' ') + ' Forms cannot go to them yet.');
    } catch (err) {
      problems.push('Row ' + r + ': ' + err.message);
      log_('Add person failed', 'Row ' + r + ': ' + err.message);
    }
  }
  ui_((done.length ? 'Ready: ' + done.join(', ') + '.' : 'Nobody new was added.') +
      (problems.length ? '\n\nNot done:\n' + problems.join('\n') : ''));
}

function person_(row) {
  var v = tab_(TABS.people).getRange(row, 1, 1, PEOPLE.length).getValues()[0];
  return {
    row: row,
    name: String(v[P['Name']] || '').trim(),
    email: String(v[P['Email']] || '').trim(),
    role: String(v[P['Role']] || '').trim(),
    guardian: String(v[P['Signs for them (guardian or representative)']] || '').trim(),
    guardianEmail: String(v[P['Their email']] || '').trim(),
    active: String(v[P['Active']] || 'Yes').trim() !== 'No',
    folder: folderId_(v[P['Folder']])
  };
}

/** Their folder, and for a worker their staff file tab. 'strict' (the menu item) also stops on a full
 * Team at a glance; sending a form only logs it, because a full team list is no reason not to send. */
function setUpPerson_(person, strict) {
  if (ROLES.indexOf(person.role) === -1) {
    throw new Error(person.name + ' needs a role: ' + ROLES.join(', ') + '.');
  }
  if (!person.folder) {
    var folder = childFolder_(recordsRoot_(), RECORDS.role[person.role]).createFolder(person.name);
    childFolder_(folder, RECORDS.sent);
    childFolder_(folder, RECORDS.done);
    if (person.role === 'Worker') childFolder_(folder, RECORDS.staff);
    tab_(TABS.people).getRange(person.row, P['Folder'] + 1).setValue(folder.getUrl());
    person.folder = folder.getId();
    log_('Folder made', person.name + ': ' + folder.getUrl());
  }
  if (person.role === 'Worker') {
    try {
      staffTab_(person);
    } catch (err) {
      if (strict) throw err;
      log_('Staff file tab not made', person.name + ': ' + err.message);
    }
  }
  return person;
}

/** A copy of the staff file tab named for the worker, and their name on Team at a glance. */
function staffTab_(person) {
  var ss = book_();
  var name = sheetName_(person.name);
  if (!ss.getSheetByName(name)) {
    var copy = tab_(TABS.staff).copyTo(ss).setName(name);
    var labels = copy.getRange(1, 1, 20, 1).getValues();
    var start = tab_(TABS.people).getRange(person.row, P['Start date'] + 1).getValue();
    var fill = { 'Worker name': person.name, 'Role': 'Worker', 'Start date': start };
    labels.forEach(function (l, i) {
      if (Object.prototype.hasOwnProperty.call(fill, l[0]) && fill[l[0]] !== '') {
        copy.getRange(i + 1, 2).setValue(fill[l[0]]);
      }
    });
    log_('Staff file tab made', name);
  }
  var team = tab_(TABS.team);
  var names = team.getRange(5, 1, 15, 1).getValues().map(function (r) { return String(r[0]).trim(); });
  if (names.indexOf(name) !== -1) return;
  var free = names.indexOf('');
  if (free === -1) {
    throw new Error('Team at a glance is full (15 workers). Add a row to it, then add ' + person.name + ' again.');
  }
  team.getRange(5 + free, 1).setValue(name);
}

/** Who signs for this person, and anything wrong with the addresses. */
function signers_(person) {
  var to = [], problems = [];
  [[person.email, person.name], [person.guardianEmail, person.guardian || 'their guardian']].forEach(function (p) {
    if (!p[0]) return;
    if (!isEmail_(p[0])) problems.push("'" + p[0] + "' for " + p[1] + ' is not an email address.');
    else to.push(p[0]);
  });
  if (!to.length && !problems.length) problems.push(person.name + ' has no email address on the People tab.');
  return { to: to, problems: problems };
}


/* ------------------------------------------------------------------ sending */

/** The forms for this person's role, grouped as on its 'Send to' tab. */
function formsFor(row) {
  var person = person_(row);
  var tabName = SEND_TABS[person.role];
  if (!tabName) throw new Error(person.name + ' needs a role on the People tab before forms can go out.');
  var templates = templates_();
  var groups = sendGroups_(tabName).map(function (g) {
    return {
      name: g.name,
      forms: g.forms.map(function (f) {
        var t = templates[f.id];
        return { id: f.id, title: f.title, need: f.need, how: t && t.ready ? ESIGN : EMAIL };
      })
    };
  });
  var who = signers_(person);
  return { person: person.name, role: person.role, to: who.to, problems: who.problems, groups: groups,
           testTo: testTo_(), from: setting_('Emails go out from') };
}

/**
 * Makes the person's copy of each chosen form, writes its row in 'Sent and returned' (not dated until
 * it really goes), and queues a test copy of every email to testTo. Nothing reaches the person here.
 */
function prepare(row, ids, testTo) {
  var person = setUpPerson_(person_(row), false);
  var who = signers_(person);
  if (who.problems.length) throw new Error(who.problems.join(' '));
  var forms = sendGroups_(SEND_TABS[person.role]).reduce(function (all, g) { return all.concat(g.forms); }, []);
  var templates = templates_();
  testTo = String(testTo || testTo_()).trim();
  if (!isEmail_(testTo)) throw new Error('Give an email address for the test copy.');
  PropertiesService.getUserProperties().setProperty('TEST_TO', testTo);  // remembered for this person
  var folder = sentFolder_(person);
  var stamp = Utilities.formatDate(new Date(), TZ, 'd MMM yyyy');
  var out = [], started = Date.now();
  ids.forEach(function (id) {
    var form = forms.filter(function (f) { return f.id === id; })[0];
    if (!form) return out.push({ title: id, error: 'That form is not on the ' + person.role + ' list.' });
    if (Date.now() - started > TIME_BUDGET_MS) return out.push({ title: form.title, error: 'Not done in time. Tick it and send it again.' });
    var name = form.title + ' - ' + person.name + ' (' + stamp + ')';
    try {
      var t = templates[id];
      var reading = readOnly_(form.need);
      var item = { title: form.title, to: who.to, how: t && t.ready && !reading ? ESIGN : EMAIL };
      if (reading) {  // they read the one document, so no copy of it is made for them
        item.copyUrl = openLink_(id);
      } else {
        var copyId = item.how === ESIGN
          ? DriveApp.getFileById(t.id).makeCopy(name, folder).getId()
          : toDoc_(id, name, folder.getId());
        item.copyUrl = DriveApp.getFileById(copyId).getUrl();
      }
      item.trackerRow = trackRow_({
        to: person.name, who: person.role, email: who.to.join(', '), form: form.title, need: form.need,
        copy: item.copyUrl, how: item.how, step: item.how === EMAIL ? 'Test' : '', testTo: testTo
      });
      if (item.how === EMAIL) item.tested = testTo;
      out.push(item);
      log_('Prepared', name + ' (' + item.how + ')');
    } catch (err) {
      log_('Prepare failed', name + ': ' + err.message);
      out.push({ title: form.title, error: err.message });
    }
  });
  return out;
}

/**
 * Queues the emails prepare() tested; the one-minute timer sends them. A row goes once its test copy
 * has gone, and only once, so pressing twice sends once.
 */
function sendNow(rows) {
  var tracker = tab_(TABS.tracker);
  return rows.map(function (r) {
    var v = tracker.getRange(r, 1, 1, TRACKER.length).getValues()[0];
    var step = v[T['Send step']];
    if (v[T['How it went out']] !== EMAIL) return { row: r, error: 'Row ' + r + ' is not an email form.' };
    if (v[T['Date sent']] || step === 'Sent') return { row: r, queued: true, note: 'Already sent.' };
    if (step === 'Send') return { row: r, queued: true, note: 'Already on its way.' };
    if (step !== 'Test sent') {
      return { row: r, error: v[T['Form']] + ': its test copy has not gone yet. Press Send again in a minute.' };
    }
    tracker.getRange(r, T['Send step'] + 1).setValue('Send');
    return { row: r, queued: true };
  });
}

/** The menu item, for when the Send panel was closed: queues the selected rows of 'Sent and returned'. */
function sendSelectedRows() {
  var sheet = SpreadsheetApp.getActiveSheet();
  if (sheet.getName() !== TABS.tracker) return ui_("Select the rows to send on 'Sent and returned' first.");
  var range = sheet.getActiveRange(), rows = [];
  for (var r = Math.max(range.getRow(), TRACKER_FIRST); r < range.getRow() + range.getNumRows(); r++) rows.push(r);
  var res = sendNow(rows);
  var bad = res.filter(function (x) { return x.error; });
  ui_((res.length - bad.length) + ' on the way; they go out within a minute.' +
      (bad.length ? '\n\nNot sent:\n' + bad.map(function (x) { return x.error; }).join('\n') : ''));
}

/**
 * Every minute, as the account that ran Set up: sends queued test copies, then queued emails. A failure
 * stays on its row as 'Not sent: why' and is told to Settings' people; nothing is retried blindly.
 */
function sendQueued(e) {
  if (!timerOwner_(e)) return;
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return;
  try {
    var tracker = tab_(TABS.tracker);
    var last = Math.max(tracker.getLastRow(), TRACKER_FIRST);
    var rows = tracker.getRange(TRACKER_FIRST, 1, last - TRACKER_FIRST + 1, TRACKER.length).getValues();
    var reading = {};  // what one person only has to read goes in one email, not one each
    var signing = {};  // and everything they have to sign goes behind one link
    rows.forEach(function (v, i) {
      var step = v[T['Send step']], r = TRACKER_FIRST + i;
      if (step !== 'Test' && step !== 'Send') return;
      if (v[T['How it went out']] === SIGN_PAGE) {
        (signing[v[T['Sent to']]] = signing[v[T['Sent to']]] || []).push({ r: r, v: v });
        return;
      }
      if (readOnly_(v[T['They need to']])) {
        var key = step + '|' + v[T['Sent to']];
        (reading[key] = reading[key] || []).push({ r: r, v: v });
        return;
      }
      var cell = tracker.getRange(r, T['Send step'] + 1);
      try {
        var person = byName_(v[T['Sent to']]);
        var copyId = fileId_(v[T['Copy sent (opens it)']]), form = { title: v[T['Form']], need: v[T['They need to']] };
        if (step === 'Test') {
          var check = email_(person, form, copyId, true);
          GmailApp.sendEmail(v[T['Test copy to']], '[Test copy for ' + person.name + '] ' + check.subject, check.body,
                             { name: setting_('Send emails as'), attachments: check.attachments });
          cell.setValue('Test sent');
          return;
        }
        var who = signers_(person);
        if (who.problems.length) throw new Error(who.problems.join(' '));
        var mail = email_(person, form, copyId, shareCopy_(copyId, who.to));
        var opts = { name: setting_('Send emails as'), attachments: mail.attachments };
        var thread = GmailApp.createDraft(who.to.join(','), mail.subject, mail.body, opts).send().getThread();
        thread.addLabel(GmailApp.getUserLabelByName(LABEL) || GmailApp.createLabel(LABEL));
        date_(tracker.getRange(r, T['Date sent'] + 1), new Date());
        tracker.getRange(r, T['Email thread'] + 1).setValue(thread.getId());
        cell.setValue('Sent');
        log_('Sent', v[T['Form']] + ' to ' + who.to.join(', '));
      } catch (err) {
        cell.setValue('Not sent: ' + err.message);
        log_('Send failed', 'Row ' + r + ': ' + err.message);
        tell_('Not sent: ' + v[T['Form']] + ' to ' + v[T['Sent to']],
              v[T['Form']] + ' to ' + v[T['Sent to']] + ' did not go out: ' + err.message + '\n\n' +
              'Fix it, then send it again from their page:\n' + personLink_(v[T['Sent to']]));
      }
    });
    Object.keys(reading).forEach(function (key) { sendReading_(tracker, reading[key]); });
    Object.keys(signing).forEach(function (name) { sendSignLink_(tracker, name, signing[name]); });
    setUpWaiting_();
    pdfsWaiting_();
  } finally {
    lock.releaseLock();
  }
}

/** Folders and staff tabs for anyone added in bulk, as many each minute as fit: a long list outlasts one call. */
function setUpWaiting_() {
  var started = Date.now(), sheet = tab_(TABS.people), last = sheet.getLastRow();
  for (var r = PEOPLE_HEAD + 1; r <= last && Date.now() - started < 120000; r++) {
    var p = person_(r), note = String(sheet.getRange(r, P['Notes'] + 1).getValue());
    if (!p.name || p.folder || note.indexOf('Not set up') === 0) continue;
    try {
      setUpPerson_(p, false);
    } catch (err) {
      sheet.getRange(r, P['Notes'] + 1).setValue('Not set up: ' + err.message);
      log_('Person not set up', p.name + ': ' + err.message);
    }
  }
}

/** One email of links for everything one person has to read, marked on every row it covers. */
function sendReading_(tracker, rows) {
  var first = rows[0].v, step = first[T['Send step']];
  var mark = function (value) {
    rows.forEach(function (x) { tracker.getRange(x.r, T['Send step'] + 1).setValue(value); });
  };
  try {
    var person = byName_(first[T['Sent to']]);
    var mail = readingEmail_(person, rows);
    var opts = { name: setting_('Send emails as') };
    if (step === 'Test') {
      GmailApp.sendEmail(first[T['Test copy to']], '[Test copy for ' + person.name + '] ' + mail.subject, mail.body, opts);
      return mark('Test sent');
    }
    var who = signers_(person);
    if (who.problems.length) throw new Error(who.problems.join(' '));
    var thread = GmailApp.createDraft(who.to.join(','), mail.subject, mail.body, opts).send().getThread();
    thread.addLabel(GmailApp.getUserLabelByName(LABEL) || GmailApp.createLabel(LABEL));
    rows.forEach(function (x) {
      date_(tracker.getRange(x.r, T['Date sent'] + 1), new Date());
      tracker.getRange(x.r, T['Email thread'] + 1).setValue(thread.getId());
    });
    mark('Sent');
    log_('Sent', rows.length + ' to read, to ' + who.to.join(', '));
  } catch (err) {
    mark('Not sent: ' + err.message);
    log_('Send failed', first[T['Sent to']] + ', ' + rows.length + ' to read: ' + err.message);
    tell_('Not sent: documents to read for ' + first[T['Sent to']], err.message + '\n\n' + personLink_(first[T['Sent to']]));
  }
}

function readingEmail_(person, rows) {
  var first = (person.guardian && !person.email ? person.guardian : person.name).split(' ')[0];
  var list = rows.map(function (x) { return '- ' + x.v[T['Form']] + '\n  ' + x.v[T['Copy sent (opens it)']]; });
  var asks = [];
  rows.forEach(function (x) {
    var need = String(x.v[T['They need to']]).replace(/\.$/, '');
    if (need !== 'Read' && asks.indexOf(need) === -1) asks.push(need);
  });
  return {
    subject: 'Harmony Care: documents to read',
    body: 'Hi ' + first + ',\n\nHarmony Care has ' + (rows.length === 1 ? 'a document' : rows.length + ' documents') +
          ' for you to read. Each link opens it.\n\n' + list.join('\n\n') + '\n\n' +
          (asks.length ? 'What to do: ' + asks.join('. ') + '.\n\n' : '') +
          'Nothing needs to come back by email. Any questions, just reply here.\n\nThank you,\nHarmony Care'
  };
}

/**
 * The timers run as one account only: the one in Settings' 'Emails go out from', which ran Set up
 * last. Any other account's timer, left over from an earlier Set up, removes itself instead of
 * sending from the wrong address or reading the wrong inbox. 'e' is the timer's event; a menu click
 * has none.
 */
function timerOwner_(e) {
  var owner = String(setting_('Emails go out from') || '').trim().toLowerCase();
  var me = Session.getEffectiveUser().getEmail().toLowerCase();
  if (!owner || owner === me) return true;
  if (e && e.triggerUid) {
    ScriptApp.getProjectTriggers().forEach(function (t) {
      if (TIMERS.indexOf(t.getHandlerFunction()) !== -1) ScriptApp.deleteTrigger(t);
    });
    log_('Timers stopped', me + ' no longer sends; ' + owner + ' does.');
  }
  return false;
}

/**
 * Everyone who can edit the sheet can send forms, so each needs the two folders the copies go in.
 * Run Set up again after adding someone to the sheet. Returns who was newly added.
 */
function shareFolders_(ss) {
  var editors = ss.getEditors().map(function (u) { return u.getEmail(); }).filter(isEmail_);
  var added = [];
  ["Folder for each person's files", 'Signing templates folder'].forEach(function (label) {
    var folder = DriveApp.getFolderById(requireFolder_(label));
    var has = folder.getEditors().map(function (u) { return u.getEmail().toLowerCase(); });
    var owner = folder.getOwner();  // none for a folder in a shared drive
    if (owner) has.push(owner.getEmail().toLowerCase());
    has.push(Session.getEffectiveUser().getEmail().toLowerCase());  // whoever runs Set up has it already
    editors.forEach(function (email) {
      if (has.indexOf(email.toLowerCase()) !== -1) return;
      try {
        folder.addEditor(email);
        if (added.indexOf(email) === -1) added.push(email);
      } catch (err) {
        log_('Folder not shared', label + ' with ' + email + ': ' + err.message);
      }
    });
  });
  return added;
}

/** The test address to offer: the last one this person used, else their own, else Settings'. */
function testTo_() {
  return PropertiesService.getUserProperties().getProperty('TEST_TO') ||
         Session.getActiveUser().getEmail() || setting_('Test copies go to');
}

/** The sidebar's 'I have requested the signature' button: dates the eSignature row. */
function markRequested(row) {
  var cell = tab_(TABS.tracker).getRange(row, T['Date sent'] + 1);
  if (!cell.getValue()) date_(cell, new Date());
  return true;
}

function email_(person, form, copyId, shared) {
  var doc = DriveApp.getFileById(copyId);
  var first = (person.guardian && !person.email ? person.guardian : person.name).split(' ')[0];
  var need = String(form.need || 'sign').replace(/\.$/, '');
  return {
    subject: 'Harmony Care: ' + form.title,
    body: 'Hi ' + first + ',\n\n' +
          'Attached is the ' + form.title + ' from Harmony Care as a PDF.' +
          (shared ? ' You can also open your own copy in Google Docs:\n' + doc.getUrl() : '') + '\n\n' +
          'What to do: ' + need + '.\n\n' +
          'When it is done, reply to this email with the signed copy attached. ' +
          'A clear photo of each signed page is fine.\n\n' +
          'Any questions, just reply here.\n\nThank you,\nHarmony Care',
    attachments: [doc.getAs('application/pdf').setName(form.title + '.pdf')]
  };
}

/** Lets the person open their own copy. An address Google cannot share with still gets the PDF. */
function shareCopy_(copyId, emails) {
  var ok = true;
  emails.forEach(function (email) {
    try {
      DriveApp.getFileById(copyId).addEditor(email);
    } catch (err) {
      ok = false;
      log_('Copy not shared', email + ': ' + err.message);
    }
  });
  return ok;
}


/* ------------------------------------------------------------------ coming back */

/** Hourly: files each signed copy that came back, marks its row, and tells Settings' people. */
function checkReturns(e) {
  if (!timerOwner_(e)) {
    if (!e) ui_('Returned forms are checked from ' + setting_('Emails go out from') +
                "'s account, where the replies arrive. It happens every hour by itself.");
    return;
  }
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return;  // the last run is still going
  try {
    var tracker = tab_(TABS.tracker);
    var last = Math.max(tracker.getLastRow(), TRACKER_FIRST);
    var rows = tracker.getRange(TRACKER_FIRST, 1, last - TRACKER_FIRST + 1, TRACKER.length).getValues();
    var claimed = {};
    rows.forEach(function (v) { var id = fileId_(v[T['Signed copy']]); if (id) claimed[id] = true; });
    var me = Session.getEffectiveUser().getEmail().toLowerCase();
    var back = 0;
    // eSignature rows waiting in the same folder share that folder's new PDFs
    var waiting = {};
    rows.forEach(function (v, i) {
      if (!v[T['Date sent']] || v[T['Returned on']] || v[T['Signed copy']] || readOnly_(v[T['They need to']])) return;
      var r = TRACKER_FIRST + i;
      try {
        var how = v[T['How it went out']];
        if ((how === EMAIL || how === SIGN_PAGE) && v[T['Email thread']]) {  // a sign page link can still get a scan by reply
          back += fromEmail_(tracker, r, v, me, claimed) ? 1 : 0;
        } else if (how === ESIGN) {
          var copy = DriveApp.getFileById(fileId_(v[T['Copy sent (opens it)']]));
          var folder = copy.getParents().next();
          (waiting[folder.getId()] = waiting[folder.getId()] || { folder: folder, rows: [] })
            .rows.push({ r: r, v: v, name: copy.getName() });
        }
      } catch (err) {
        log_('Check failed', 'Row ' + r + ': ' + err.message);
      }
    });
    Object.keys(waiting).forEach(function (id) {
      back += fromFolder_(tracker, waiting[id], claimed);
    });
    if (back) log_('Came back', back + ' form' + (back > 1 ? 's' : ''));
  } finally {
    lock.releaseLock();
  }
}

function fromEmail_(tracker, r, v, me, claimed) {
  var thread = GmailApp.getThreadById(v[T['Email thread']]);
  if (!thread) return false;
  var sent = v[T['Date sent']];
  var replies = thread.getMessages().filter(function (m) {
    return m.getFrom().toLowerCase().indexOf(me) === -1 && day_(m.getDate()) >= day_(sent);
  });
  if (!replies.length) return false;
  var files = [];
  replies.forEach(function (m) {
    m.getAttachments({ includeInlineImages: false }).forEach(function (a) { files.push(a); });
  });
  var notes = String(v[T['Notes']] || '');
  if (!files.length) {
    if (notes.indexOf('Replied without') === -1) {
      tracker.getRange(r, T['Notes'] + 1).setValue(
        (notes ? notes + ' ' : '') + 'Replied without the signed form attached (' + day_(new Date()) + ').');
      var said = replies.map(function (m) { return String(m.getPlainBody() || '').trim().slice(0, 600); })
        .filter(String).join('\n---\n');
      tell_('Reply without the form: ' + v[T['Form']] + ' from ' + v[T['Sent to']],
            v[T['Sent to']] + ' replied about ' + v[T['Form']] + ' but did not attach the signed form. ' +
            'They wrote:\n\n' + (said || '(nothing)') + '\n\n' +
            'Reply to them at ' + v[T['Email']] + ', or ask ' + setting_('Emails go out from') +
            ' to answer in the thread (Gmail label: ' + LABEL + ').');
    }
    return false;
  }
  var folder = doneFolder_(byName_(v[T['Sent to']]));
  var saved = files.map(function (a, n) {
    var ext = (a.getName().match(/\.[A-Za-z0-9]+$/) || [''])[0];
    var name = v[T['Form']] + ' - ' + v[T['Sent to']] + ' - signed' + (files.length > 1 ? ' ' + (n + 1) : '') + ext;
    return folder.createFile(a.copyBlob().setName(name))
      .setDescription(MARK + 'signed copy from email, Sent and returned row ' + r);
  });
  saved.forEach(function (f) {
    claimed[f.getId()] = true;
    listCompleted_(f);
  });
  returned_(tracker, r, v, saved[0].getUrl(), saved.length > 1 ? saved.length + ' files, all in their folder.' : '');
  return true;
}

/**
 * Google files a completed eSignature PDF beside the document it was requested from. A new PDF in
 * the folder whose name holds the copy's name is that copy's. A new PDF that matches no copy goes to
 * the only waiting row if there is one, and otherwise to a person to match by hand.
 */
function fromFolder_(tracker, group, claimed) {
  var since = group.rows.map(function (w) { return day_(w.v[T['Date sent']]); }).sort()[0];
  var fresh = [];
  var it = group.folder.getFilesByType('application/pdf');
  while (it.hasNext()) {
    var f = it.next();
    if (claimed[f.getId()] || (f.getDescription() || '').indexOf(MARK) === 0) continue;
    if (day_(f.getDateCreated()) >= since) fresh.push(f);
  }
  var n = 0;
  var open = group.rows.slice();
  fresh.forEach(function (f) {
    var match = open.filter(function (w) {
      return f.getName().toLowerCase().indexOf(w.name.toLowerCase()) !== -1;
    })[0];
    if (!match && open.length === 1) match = open[0];
    if (!match) {
      claimed[f.getId()] = true;
      f.setDescription(MARK + 'signed file not matched to a form; match it by hand');  // asked once, not hourly
      log_('Unmatched signed file', f.getName() + ' ' + f.getUrl());
      tell_('A signed file to match: ' + f.getName(),
            'A new signed PDF arrived in ' + group.folder.getName() + ' and more than one form there is ' +
            'waiting, so it was not matched.\n\nOpen it: ' + f.getUrl() + '\n\n' +
            'Then press Mark as back on the right form on their page:\n' + TOOL_URL + '#forms');
      return;
    }
    claimed[f.getId()] = true;
    open.splice(open.indexOf(match), 1);
    f.moveTo(doneFolder_(byName_(match.v[T['Sent to']])));  // Google files it beside the copy, in Sent out
    listCompleted_(f);
    returned_(tracker, match.r, match.v, f.getUrl(), '');
    n++;
  });
  return n;
}

function returned_(tracker, r, v, url, note) {
  date_(tracker.getRange(r, T['Returned on'] + 1), new Date());
  tracker.getRange(r, T['Signed copy'] + 1).setValue(url);
  if (note) {
    var notes = String(v[T['Notes']] || '');
    tracker.getRange(r, T['Notes'] + 1).setValue((notes ? notes + ' ' : '') + note);
  }
  tell_('Signed and back: ' + v[T['Form']] + ' from ' + v[T['Sent to']],
        v[T['Sent to']] + ' has returned ' + v[T['Form']] + '.\n\nOpen the signed copy: ' + url + '\n\n' +
        'Check it, then approve it here:\n' + personLink_(v[T['Sent to']]));
}


/* ------------------------------------------------------------------ chasing and dates */

/**
 * Daily: a form out longer than 'Chase after (days)' gets a reminder drafted in its Gmail thread (never
 * sent by the script), at most once a week, and one email lists them all.
 */
function dailyChase(e) {
  if (!timerOwner_(e)) return;
  var tracker = tab_(TABS.tracker);
  var last = Math.max(tracker.getLastRow(), TRACKER_FIRST);
  var rows = tracker.getRange(TRACKER_FIRST, 1, last - TRACKER_FIRST + 1, TRACKER.length).getValues();
  var days = Number(setting_('Chase after (days)')) || 14;
  var today = day_(new Date());
  var lines = [];
  rows.forEach(function (v, i) {
    var sent = v[T['Date sent']], r = TRACKER_FIRST + i;
    if (!sent || v[T['Returned on']] || !v[T['Form']] || readOnly_(v[T['They need to']])) return;
    if (daysBetween_(day_(sent), today) < days) return;
    var chased = v[T['Last chased']];
    if (chased && daysBetween_(day_(chased), today) < 7) return;
    try {
      var how = v[T['How it went out']];
      if ((how === EMAIL || how === SIGN_PAGE) && v[T['Email thread']]) {
        var thread = GmailApp.getThreadById(v[T['Email thread']]);
        if (thread) {
          thread.createDraftReply('Hi, just a reminder about the ' + v[T['Form']] + ' we sent on ' +
            pretty_(sent) + (how === SIGN_PAGE
              ? '. The link in the email below still opens it, when you have a moment.'
              : '. Could you sign it and reply with the signed copy when you have a moment?') +
            '\n\nThank you,\nHarmony Care');
        }
        lines.push('- ' + v[T['Form']] + ', ' + v[T['Sent to']] + ': a reminder is waiting in Gmail drafts.');
      } else if (how === BY_HAND) {
        lines.push('- ' + v[T['Form']] + ', ' + v[T['Sent to']] + ': handed over, ask for it back in person.');
      } else {
        lines.push('- ' + v[T['Form']] + ', ' + v[T['Sent to']] + ': open the copy and send a reminder from ' +
                   'its eSignature panel. ' + v[T['Copy sent (opens it)']]);
      }
      date_(tracker.getRange(r, T['Last chased'] + 1), new Date());
    } catch (err) {
      log_('Chase failed', 'Row ' + r + ': ' + err.message);
    }
  });
  if (lines.length) {
    tell_(lines.length + ' form' + (lines.length > 1 ? 's' : '') + ' to chase',
          'Out for ' + days + ' days or more and not back yet:\n\n' + lines.join('\n') + '\n\n' + TOOL_URL + '#forms');
    log_('Chased', lines.length + ' forms');
  }
}

/** Monday 8 am, or from the menu: what is overdue or due soon on the Everything tab, and forms out. */
function mondayDigest(e) {
  if (e && !timerOwner_(e)) return;  // from the menu anyone may ask for it
  var now = [], soon = [];
  checklist_().modules.forEach(function (m) {
    m.standards.forEach(function (st) {
      st.items.forEach(function (item) {
        var line = '- ' + item.title + ' (' + st.name + ')' + (item.due ? ', due ' + item.due : '') + ': ' + item.status;
        if (/^(Overdue|Missing|Expired)/.test(item.status)) now.push(line);
        else if (/^(Due within|Renew within)/.test(item.status)) soon.push(line);
      });
    });
  });
  var tracker = tab_(TABS.tracker);
  var last = Math.max(tracker.getLastRow(), TRACKER_FIRST);
  var out = tracker.getRange(TRACKER_FIRST, 1, last - TRACKER_FIRST + 1, TRACKER.length).getValues()
    .filter(function (v) { return v[T['Date sent']] && !v[T['Returned on']] && !readOnly_(v[T['They need to']]); }).length;
  var approve = tracker.getRange(TRACKER_FIRST, 1, last - TRACKER_FIRST + 1, TRACKER.length).getValues()
    .filter(function (v) { return v[T['Returned on']] && !v[T['Approved on']]; }).length;
  tell_('Harmony Care this week: ' + now.length + ' overdue or missing, ' + soon.length + ' due soon',
        'Overdue or missing (' + now.length + '):\n' + (now.join('\n') || '- nothing') + '\n\n' +
        'Due in the next 30 days (' + soon.length + '):\n' + (soon.join('\n') || '- nothing') + '\n\n' +
        'Forms out and not back: ' + out + '. Back and waiting for approval: ' + approve + '.\n\n' + TOOL_URL);
}

/** The menu item: today's date in 'Last reviewed' for the document on the selected row. */
function markReviewedToday() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var row = sheet.getActiveRange().getRow();
  var no = sheet.getName() === TABS.checklist ? sheet.getRange(row, 1).getValue()
    : sheet.getName() === TABS.everything ? numberOf_(sheet.getRange(row, 4).getValue() || sheet.getRange(row, 5).getValue()).no : '';
  if (typeof no !== 'number') return ui_('Select a document row on the Everything or Checklist tab.');
  var title = reviewed_(no);
  book_().toast('No. ' + no + ' marked reviewed today. Next due and Status update by themselves.', 'Harmony forms');
  return title;
}

/** Today's date in 'Last reviewed' for a document; Next due and Status follow from it. */
function reviewed_(no) {
  var checklist = tab_(TABS.checklist);
  var head = checklist.getRange(5, 1, 1, checklist.getLastColumn()).getValues()[0];
  var lastCol = head.indexOf('Last reviewed (type the date)') + 1;
  if (!lastCol) throw new Error("The Checklist tab has changed: no 'Last reviewed (type the date)' column.");
  var nos = checklist.getRange(1, 1, checklist.getLastRow(), 1).getValues(), target = 0;
  for (var i = 0; i < nos.length; i++) if (nos[i][0] === no) { target = i + 1; break; }
  if (!target) throw new Error('Document No. ' + no + ' is not on the Checklist tab.');
  date_(checklist.getRange(target, lastCol), new Date());
  var title = checklist.getRange(target, 4).getValue();
  log_('Reviewed', 'No. ' + no + ', ' + title);
  return title;
}

/** '12. Conflict of Interest Register' as its number and its name. */
function numberOf_(label) {
  var m = String(label || '').match(/^(\d+)\.\s+(.*)$/);
  return m ? { no: Number(m[1]), title: m[2] } : { no: '', title: String(label || '') };
}


/* ------------------------------------------------------------------ bringing the checklist up to date */

// An older copy of the checklist has no Everything tab and three wordings since corrected. Rather
// than swap the sheet for a new file, upgrade_ makes any copy current from its own tabs. Each part
// looks before it changes anything, so running it twice changes nothing the second time.

var WORDING = [  // [tab, old text, new text]
  ['Start here', 'Laid out like the Avaana checklist: module, Practice Standard, its outcome and indicators, ' +
   'then every document for it.', 'The documents on their own, by module and Practice Standard with its ' +
   "outcome and indicators, and each document's version, approval and review dates."],
  ['Staff file (copy per worker)', ' Avaana checklist: auditors expect a formal agreement for each worker.', ''],
  ['Staff file (copy per worker)', 'Avaana registration checklist (key management personnel). ', '']
];
/** The menu item. Anyone who can edit the sheet may run it; it sends nothing. */
function upgradeWorkbook() {
  var done = upgrade_(book_());
  ui_(done.length ? 'Brought up to date:\n- ' + done.join('\n- ') : 'The checklist was already up to date.');
}

function upgrade_(ss) {
  var done = [];
  WORDING.forEach(function (w) {
    var n = tab_(w[0]).createTextFinder(w[1]).replaceAllWith(w[2]);
    if (n) done.push(w[0] + ': ' + (w[2] ? 'reworded' : 'removed') + ' "' + w[1].trim().slice(0, 40) + '..."');
  });
  // A blank staff file row was labelled Missing; nobody has looked in the worker's file, so it is
  // only not recorded yet.
  var staff = tab_(TABS.staff), team = tab_(TABS.team);
  var n = staff.createTextFinder('"Missing"').matchFormulaText(true).replaceAllWith('"Not recorded yet"');
  n += team.createTextFinder('"Missing"').matchFormulaText(true).replaceAllWith('"Not recorded yet"');
  if (team.getRange('B4').getValue() === 'Missing or expired') team.getRange('B4').setValue('Not recorded or expired');
  var rules = staff.getConditionalFormatRules(), changed = false;
  rules = rules.map(function (rule) {
    var c = rule.getBooleanCondition();
    var f = c && c.getCriteriaValues()[0];
    if (typeof f !== 'string' || !/="Missing"$/.test(f)) return rule;
    changed = true;
    return SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(f.replace('"Missing"', '"Not recorded yet"'))
      .setBackground('#FCE8B2').setRanges(rule.getRanges()).build();
  });
  if (changed) staff.setConditionalFormatRules(rules);
  if (n || changed) done.push("Staff file: 'Missing' is now 'Not recorded yet'");
  var tracker = tab_(TABS.tracker);
  var status = tracker.getRange(TRACKER_FIRST, T['Status'] + 1);
  if (String(status.getFormula()).indexOf('nothing to return') === -1) {
    var last = tracker.getMaxRows(), lines = [];
    for (var r = TRACKER_FIRST; r <= last; r++) lines.push([statusFormula_(r)]);
    tracker.getRange(TRACKER_FIRST, T['Status'] + 1, lines.length, 1).setFormulas(lines);
    done.push("Sent and returned: what is only read no longer says chase up");
  }
  var left = ss.createTextFinder('avaana').findAll().length + ss.createTextFinder('havana').findAll().length;
  if (left) log_('Still named', left + ' cell(s) still name the 2023 consultancy');
  if (done.length) log_('Brought up to date', done.join('; '));
  return done;
}

/** The status line of a 'Sent and returned' row. */
function statusFormula_(r) {
  return '=IF(E' + r + '="","",IF(I' + r + '<>"","\u2713 Done and approved",IF(G' + r + '<>"","Returned: check and approve",' +
    'IF(A' + r + '="","Not sent yet",IF(AND(LEFT(F' + r + ',4)="Read",LEFT(F' + r + ',13)<>"Read and sign"),"Sent, nothing to return",' +
    'IF(AND(ISNUMBER(A' + r + '),A' + r + '<TODAY()-14),"Chase up: sent over 2 weeks ago","Waiting for it to come back"))))))';
}


/* ------------------------------------------------------------------ signing templates */

/** A Google Doc of every form on the three 'Send to' tabs, once each, listed on Signing templates. */
function makeTemplates() {
  var started = Date.now();
  var folderId = requireFolder_('Signing templates folder');
  var sheet = tab_(TABS.templates);
  var have = templates_();
  var forms = {};
  Object.keys(SEND_TABS).forEach(function (role) {
    sendGroups_(SEND_TABS[role]).forEach(function (g) {
      g.forms.forEach(function (f) {
        (forms[f.id] = forms[f.id] || { title: f.title, roles: [] }).roles.push(role);
      });
    });
  });
  var made = 0, left = 0;
  Object.keys(forms).forEach(function (id) {
    if (have[id]) return;
    if (Date.now() - started > TIME_BUDGET_MS) { left++; return; }
    var f = forms[id];
    try {
      var docId = toDoc_(id, f.title, folderId);
      var r = Math.max(sheet.getLastRow(), TEMPLATE_HEAD) + 1;
      sheet.getRange(r, 1, 1, TEMPLATES.length).setValues([[
        f.title, f.roles.join(', '), DriveApp.getFileById(docId).getUrl(), false,
        'https://drive.google.com/file/d/' + id + '/view']]);
      sheet.getRange(r, TP['Signature fields placed'] + 1).insertCheckboxes();
      made++;
    } catch (err) {
      log_('Template failed', f.title + ': ' + err.message);
    }
  });
  log_('Signing templates', made + ' made' + (left ? ', ' + left + ' still to make' : ''));
  ui_(made + ' signing template' + (made === 1 ? '' : 's') + ' made.' +
      (left ? ' ' + left + ' still to make: run Make signing templates again.' : '') +
      '\n\nOpen each one, place the eSignature fields, and tick its box on the Signing templates tab.');
}

/** {original file id: {id: template doc id, ready: fields placed}} */
function templates_() {
  var sheet = tab_(TABS.templates);
  var last = sheet.getLastRow();
  var out = {};
  if (last <= TEMPLATE_HEAD) return out;
  sheet.getRange(TEMPLATE_HEAD + 1, 1, last - TEMPLATE_HEAD, TEMPLATES.length).getValues().forEach(function (v) {
    var original = fileId_(v[TP['Original']]), doc = fileId_(v[TP['Signing template (open it to place the fields)']]);
    if (original && doc) out[original] = { id: doc, ready: v[TP['Signature fields placed']] === true };
  });
  return out;
}


/* ------------------------------------------------------------------ the workbook's own tabs */

/** [{name, forms: [{id, title, need}]}] from a 'Send to' tab, the way the tab groups them. */
function sendGroups_(name) {
  var sheet = tab_(name);
  var head = 5;
  var last = sheet.getLastRow();
  if (last <= head) return [];
  var range = sheet.getRange(head + 1, 1, last - head, 6);
  var values = range.getValues(), links = range.getRichTextValues(), formulas = range.getFormulas();
  var groups = [];
  values.forEach(function (v, i) {
    if (typeof v[0] !== 'number' && v[0] && !v[1]) return groups.push({ name: String(v[0]), forms: [] });
    if (typeof v[0] !== 'number' || !groups.length) return;
    var url = links[i][1] && links[i][1].getLinkUrl();
    if (!url) url = (String(formulas[i][1]).match(/HYPERLINK\("([^"]+)"/i) || [])[1];
    var id = fileId_(url);
    if (id) {
      groups[groups.length - 1].forms.push({ id: id, title: String(v[1]), need: String(v[2] || ''), when: String(v[3] || ''),
                                             fills: String(v[4] || ''), signs: String(v[5] || '') });
    }
  });
  return groups.filter(function (g) { return g.forms.length; });
}

/** A document that is only read: nothing comes back to chase. 'Read and sign' does come back. */
function readOnly_(need) {
  var n = String(need || '');
  return /^Read\b/.test(n) && !/^Read and sign/.test(n);
}

/** Writes a new row in 'Sent and returned' (left undated) and returns its row number. */
function trackRow_(x) {
  var tracker = tab_(TABS.tracker);
  var last = Math.max(tracker.getLastRow(), TRACKER_FIRST);
  var forms = tracker.getRange(TRACKER_FIRST, T['Form'] + 1, last - TRACKER_FIRST + 1, 1).getValues();
  var r = TRACKER_FIRST + forms.length;
  for (var i = 0; i < forms.length; i++) if (!forms[i][0]) { r = TRACKER_FIRST + i; break; }
  tracker.getRange(r, T['Sent to'] + 1, 1, 5).setValues([[x.to, x.who, x.email, x.form, x.need]]);
  tracker.getRange(r, T['Copy sent (opens it)'] + 1, 1, 2).setValues([[x.copy, x.how]]);
  tracker.getRange(r, T['Send step'] + 1, 1, 2).setValues([[x.step, x.step ? x.testTo : '']]);
  return r;
}


/* ------------------------------------------------------------------ Drive */

/** Copies a Word or PDF form into a Google Doc. Drive's own API does the conversion. */
function toDoc_(fileId, name, folderId) {
  var res = UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + fileId +
                              '/copy?supportsAllDrives=true&fields=id', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    payload: JSON.stringify({ name: name, mimeType: 'application/vnd.google-apps.document', parents: [folderId] })
  });
  if (res.getResponseCode() !== 200) {
    throw new Error('Google could not make a Google Doc of ' + name + ' (' + res.getResponseCode() + ': ' +
                    res.getContentText().slice(0, 200) + ')');
  }
  return JSON.parse(res.getContentText()).id;
}

/** Which of the linked files open in Docs or Sheets, kept so the tool need not ask Drive for each on every read. */
function learnKinds_() {
  var kinds = {};
  ['Send to participants', 'Send to workers', 'Send to practitioners', 'Harmony fills in', TABS.everything, TABS.checklist, 'Year planner', TABS.staff].forEach(function (name) {
    var sheet = book_().getSheetByName(name);
    if (!sheet) return;
    sheet.getDataRange().getRichTextValues().forEach(function (row) {
      row.forEach(function (cell) {
        var id = cell && fileId_(cell.getLinkUrl());
        if (!id || id in kinds) return;
        try {
          var mime = DriveApp.getFileById(id).getMimeType();
          kinds[id] = /sheet|excel/.test(mime) ? 's' : /word|document/.test(mime) ? 'd' : '';  // an xlsx's type says officedocument
        } catch (err) {
          log_('File not read', id + ': ' + err.message);
        }
      });
    });
  });
  var keep = {};
  Object.keys(kinds).forEach(function (id) { if (kinds[id]) keep[id] = kinds[id]; });
  PropertiesService.getScriptProperties().setProperty('KINDS', JSON.stringify(keep));
  return Object.keys(keep).length;
}

/** The address that opens a file in Google Docs or Sheets, not as a Word download. */
function openLink_(idOrUrl) {
  var id = /^https:/.test(String(idOrUrl)) ? fileId_(idOrUrl) : String(idOrUrl);
  if (!id) return /^https:/.test(String(idOrUrl)) ? idOrUrl : '';
  var kinds = openLink_.kinds || (openLink_.kinds = JSON.parse(PropertiesService.getScriptProperties().getProperty('KINDS') || '{}'));
  if (kinds[id] === 'd') return 'https://docs.google.com/document/d/' + id + '/edit';
  if (kinds[id] === 's') return 'https://docs.google.com/spreadsheets/d/' + id + '/edit';
  return 'https://drive.google.com/file/d/' + id + '/view';
}

function fileId_(url) {
  var m = String(url || '').match(/\/d\/([\w-]{20,})|[?&]id=([\w-]{20,})/);
  return m ? (m[1] || m[2]) : '';
}

/** A folder inside another, made the first time it is asked for. */
function childFolder_(parent, name) {
  var found = parent.getFoldersByName(name);
  return found.hasNext() ? found.next() : parent.createFolder(name);
}

function recordsRoot_() {
  return DriveApp.getFolderById(requireFolder_("Folder for each person's files"));
}

function allCompleted_() {
  return childFolder_(recordsRoot_(), RECORDS.all);
}

function personFolder_(person) {
  return DriveApp.getFolderById(setUpPerson_(person, false).folder);
}

function sentFolder_(person) {
  return childFolder_(personFolder_(person), RECORDS.sent);
}

function doneFolder_(person) {
  return childFolder_(personFolder_(person), RECORDS.done);
}

/** A shortcut, so the one signed copy lives in the person's folder and shows in the list of everything. */
function listCompleted_(file) {
  try {
    allCompleted_().createShortcut(file.getId()).setName(file.getName());
  } catch (err) {
    log_('Not listed', file.getName() + ': ' + err.message);  // the file is safe in their folder either way
  }
}

/** Someone who signs in to the tool needs the signed copies it links to. */
function shareRecords_(email) {
  try {
    var root = recordsRoot_();
    var has = root.getEditors().map(function (u) { return u.getEmail().toLowerCase(); });
    var owner = root.getOwner();
    if (owner) has.push(owner.getEmail().toLowerCase());
    if (has.indexOf(email) === -1) root.addEditor(email);
  } catch (err) {
    log_('Records not shared', email + ': ' + err.message);
  }
}

function folderId_(url) {
  var m = String(url || '').match(/\/folders\/([\w-]{20,})/);
  return m ? m[1] : '';
}

function requireFolder_(label) {
  var id = folderId_(setting_(label));
  if (!id) throw new Error("Settings: '" + label + "' needs a Drive folder link. Run Harmony forms > Set up.");
  return id;
}


/* ------------------------------------------------------------------ small things */

/** The Sheet, by the id Set up saved: a call from the tool has no open Sheet to ask. */
function book_() {
  var id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  return id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActive();
}

function tab_(name) {
  var sheet = book_().getSheetByName(name);
  if (!sheet) throw new Error("There is no '" + name + "' tab. Run Harmony forms > Set up, or check its name.");
  return sheet;
}

function ensureTab_(ss, name, title, intro, headRow, head) {
  var sheet = ss.getSheetByName(name);
  if (sheet) return sheet;
  sheet = ss.insertSheet(name);
  sheet.getRange(1, 1).setValue(title).setFontSize(14).setFontWeight('bold').setFontColor('#38553E');
  sheet.getRange(2, 1).setValue(intro).setFontColor('#5F564C');
  sheet.getRange(headRow, 1, 1, head.length).setValues([head])
    .setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#38553E').setWrap(true);
  sheet.setFrozenRows(headRow);
  return sheet;
}

function setting_(label) {
  var sheet = book_().getSheetByName(TABS.settings);
  if (!sheet || sheet.getLastRow() < 4) return '';
  var rows = sheet.getRange(4, 1, sheet.getLastRow() - 3, 2).getValues();
  for (var i = 0; i < rows.length; i++) if (rows[i][0] === label) return rows[i][1];
  return '';
}

function putSetting_(label, value) {
  var sheet = tab_(TABS.settings);
  var rows = sheet.getRange(4, 1, Math.max(sheet.getLastRow() - 3, 1), 1).getValues();
  for (var i = 0; i < rows.length; i++) {
    if (rows[i][0] === label) return sheet.getRange(4 + i, 2).setValue(value);
  }
  throw new Error("Settings has no '" + label + "' row.");
}

/** Everyone on the Settings line hears about it; a bad address there is logged, never fatal. */
function tell_(subject, body) {
  var to = String(setting_('Tell these people when a form comes back') || '')
    .split(',').map(function (s) { return s.trim(); }).filter(isEmail_);
  if (!to.length) return log_('Not told', subject + ' (no address in Settings)');
  try {
    MailApp.sendEmail({ to: to.join(','), subject: subject, body: body, name: setting_('Send emails as') });
  } catch (err) {
    log_('Notice failed', subject + ': ' + err.message);
  }
}

/** The address of one person's page in the tool, or the tool if they have since left the list. */
function personLink_(name) {
  try {
    return TOOL_URL + '#person/' + byName_(name).row;
  } catch (err) {
    return TOOL_URL + '#forms';
  }
}

function log_(what, detail) {
  var sheet = book_().getSheetByName(TABS.log);
  if (sheet) sheet.appendRow([Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm'), what, detail]);
}

function ui_(message) {
  try { SpreadsheetApp.getUi().alert(message); } catch (err) { log_('Message', message); }  // timers have no screen
}

function byName_(name) {
  var sheet = tab_(TABS.people);
  var names = sheet.getRange(PEOPLE_HEAD + 1, 1, Math.max(sheet.getLastRow() - PEOPLE_HEAD, 1), 1).getValues();
  for (var i = 0; i < names.length; i++) {
    if (String(names[i][0]).trim() === String(name).trim()) return person_(PEOPLE_HEAD + 1 + i);
  }
  throw new Error(name + ' is no longer on the People tab.');
}

/** Dates go in as yyyy-mm-dd text, which Sheets reads as that day whatever the script's time zone. */
function date_(range, d) {
  range.setValue(day_(d)).setNumberFormat('d mmm yyyy');
}

function day_(d) {
  return d instanceof Date ? Utilities.formatDate(d, TZ, 'yyyy-MM-dd') : String(d || '').slice(0, 10);
}

function pretty_(d) {
  var key = day_(d);
  return Utilities.formatDate(new Date(key + 'T12:00:00Z'), 'UTC', 'd MMMM');
}

function daysBetween_(a, b) {
  return Math.round((Date.parse(b + 'T00:00:00Z') - Date.parse(a + 'T00:00:00Z')) / 86400000);
}

function isEmail_(s) {
  return /^[^\s@,;]+@[^\s@,;]+\.[^\s@,;]+$/.test(String(s || '').trim());
}

/** A tab name Team at a glance can find: no characters Sheets refuses, no straight apostrophe. */
function sheetName_(name) {
  return String(name).replace(/[\[\]\*\?:\/\\]/g, ' ').replace(/'/g, '’').replace(/\s+/g, ' ').trim()
    .slice(0, 90);
}

function index_(names) {
  var out = {};
  names.forEach(function (n, i) { out[n] = i; });
  return out;
}


/* ------------------------------------------------------------------ the tool's way in */

// hc.riseos.care/tool holds no data. It posts here, as a web app that runs as the Set up account, so the
// people using it need no access to the Sheet and never meet a Google permission screen. Anyone on the
// internet can reach doPost, so nothing past askCode_ and signIn_ happens without a live session, and
// no page is served from this script: the only way in is the one door below.

var CODE_MINUTES = 10, CODE_TRIES = 5, CODES_AN_HOUR = 5, SESSION_DAYS = 14;

function doPost(e) {
  var out;
  try {
    out = route_(JSON.parse(e.postData.contents));
  } catch (err) {
    log_('Tool call failed', String(err.message).slice(0, 300));
    out = { ok: false, error: err.said ? err.message : 'That did not work. Try again in a moment.' };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function route_(body) {
  var action = String(body.action || '');
  if (action === 'code') return askCode_(body);
  if (action === 'signin') return signIn_(body);
  if (Object.prototype.hasOwnProperty.call(SIGNING, action)) return SIGNING[action](body);  // holds its own key
  var who = session_(body.token);
  if (!who) return { ok: false, signedOut: true, error: 'Please sign in again.' };
  if (!Object.prototype.hasOwnProperty.call(TOOL, action)) throw said_('The tool asked for something this script does not do.');
  try {
    return TOOL[action](body, who);
  } catch (err) {  // past the door the person is staff, and this script's errors are written for them
    log_('Tool said', who + ', ' + action + ': ' + String(err.message).slice(0, 300));
    return { ok: false, error: err.message };
  }
}

function said_(message) {
  var err = new Error(message);
  err.said = true;
  return err;
}

function hash_(text) {
  return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(text)));
}

function emailKey_(email) {
  return hash_(String(email || '').trim().toLowerCase());
}

/** Answers the same whether or not the address may sign in, so the form cannot be used to test addresses. */
function askCode_(body) {
  var email = String(body.email || '').trim().toLowerCase();
  var allowed = String(setting_('Who can sign in to the tool') || '').toLowerCase().split(',')
    .map(function (s) { return s.trim(); });
  if (!isEmail_(email) || allowed.indexOf(email) === -1) return { ok: true };
  var cache = CacheService.getScriptCache(), key = emailKey_(email);
  var sent = Number(cache.get('asked:' + key)) || 0;
  if (sent >= CODES_AN_HOUR) return { ok: true };
  cache.put('asked:' + key, String(sent + 1), 3600);
  var digits = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, Utilities.getUuid());
  var n = digits.slice(0, 6).reduce(function (all, b) { return (all * 256 + (b & 255)) % 1000000; }, 0);
  var code = ('000000' + n).slice(-6);
  cache.put('code:' + key, JSON.stringify({ hash: hash_(key + code), tries: 0 }), CODE_MINUTES * 60);
  MailApp.sendEmail({ to: email, subject: 'Your Harmony Care sign-in code: ' + code, name: setting_('Send emails as'),
                      body: 'Your code is ' + code + '. It works for ' + CODE_MINUTES + ' minutes.\n\n' +
                            'If you did not ask for it, you can ignore this email.' });
  return { ok: true };
}

function signIn_(body) {
  var email = String(body.email || '').trim().toLowerCase(), code = String(body.code || '').replace(/\D/g, '');
  var cache = CacheService.getScriptCache(), key = emailKey_(email);
  var held = cache.get('code:' + key);
  if (!held) return { ok: false, error: 'That code has run out. Ask for a new one.' };
  held = JSON.parse(held);
  if (hash_(key + code) !== held.hash) {
    held.tries++;
    if (held.tries >= CODE_TRIES) {
      cache.remove('code:' + key);
      return { ok: false, error: 'Too many tries. Ask for a new code.' };
    }
    cache.put('code:' + key, JSON.stringify(held), CODE_MINUTES * 60);
    return { ok: false, error: "That code isn't right. Check it and try again." };
  }
  cache.remove('code:' + key);
  var token = (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '');
  var props = PropertiesService.getScriptProperties(), now = Date.now();
  var all = props.getProperties();
  Object.keys(all).forEach(function (k) {  // a session that ran out stops taking up room here
    if (k.indexOf('session:') === 0 && JSON.parse(all[k]).until < now) props.deleteProperty(k);
  });
  props.setProperty('session:' + hash_(token), JSON.stringify({ email: email, until: now + SESSION_DAYS * 86400000 }));
  shareRecords_(email);
  log_('Signed in', email);
  return { ok: true, token: token, email: email };
}

/** Who the token belongs to, or nothing. Only its hash is stored, so the Sheet's owner never holds a token. */
function session_(token) {
  if (typeof token !== 'string' || token.length < 32) return '';
  var held = PropertiesService.getScriptProperties().getProperty('session:' + hash_(token));
  if (!held) return '';
  held = JSON.parse(held);
  if (held.until < Date.now()) return '';
  // taken off the list means out, whatever the session says
  var allowed = String(setting_('Who can sign in to the tool') || '').toLowerCase().split(',')
    .map(function (s) { return s.trim(); });
  return allowed.indexOf(held.email) === -1 ? '' : held.email;
}

/** One change at a time, so two people saving together cannot land on the same row. */
function locked_(fn) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) throw said_('Someone else is saving. Try again in a moment.');
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

function theRow_(value) {
  var r = Number(value);
  if (!Number.isInteger(r) || r < TRACKER_FIRST) throw said_('That row is not on the list.');
  if (!tab_(TABS.tracker).getRange(r, T['Form'] + 1).getValue()) throw said_('That row is not on the list.');
  return r;
}

function theIds_(ids) {
  if (!Array.isArray(ids) || !ids.length) throw said_('Tick at least one form first.');
  ids.forEach(function (id) { if (!/^[\w-]{20,}$/.test(String(id))) throw said_('That form is not on the list.'); });
  return ids;
}

/** Words typed in the tool, safe to keep in a cell: a leading = + - @ would otherwise run as a formula. */
function words_(value, most) {
  var text = String(value == null ? '' : value).trim().slice(0, most);
  return /^[=+\-@]/.test(text) ? "'" + text : text;
}

var TOOL = {
  load: function () {
    return { ok: true, today: day_(new Date()), checklist: checklist_(), forms: forms_(), notes: startNotes_() };
  },
  book: function () { return { ok: true, data: { office: listOf_('Harmony fills in'), planner: planner_(), rules: rules_() } }; },
  team: function () { return { ok: true, data: team_() }; },
  staffFile: function (b) { return { ok: true, data: staffFile_(theWorker_(b.person)) }; },
  staffDate: function (b, who) {
    var worker = theWorker_(b.person);
    locked_(function () { staffDate_(worker, b, who); });
    SpreadsheetApp.flush();
    return { ok: true, data: staffFile_(worker) };
  },
  staffFlag: function (b, who) {
    var worker = theWorker_(b.person), row = Number(b.sheetRow);
    if (!Number.isInteger(row) || row < 9 || row > 19) throw said_('That is not on their file.');
    locked_(function () {
      book_().getSheetByName(sheetName_(worker.name)).getRange(row, 2).setValue(b.yes ? 'Yes' : 'No');
      log_('Staff file', worker.name + ', row ' + row + ' by ' + who);
    });
    SpreadsheetApp.flush();
    return { ok: true, data: staffFile_(worker) };
  },
  plannerDone: function (b, who) {
    var r = Number(b.row), sheet = tab_('Year planner');
    if (!Number.isInteger(r) || r < 4 || r > sheet.getLastRow() || !sheet.getRange(r, 2).getValue()) throw said_('That job is not on the planner.');
    locked_(function () {
      var cell = sheet.getRange(r, 5);
      if (b.undo) cell.setValue(''); else date_(cell, new Date());
      log_('Year planner', 'Row ' + r + (b.undo ? ' cleared' : ' done') + ' by ' + who);
    });
    SpreadsheetApp.flush();
    return { ok: true, today: day_(new Date()), checklist: checklist_(), data: { planner: planner_() } };
  },
  people: function (b, who) {
    var result = locked_(function () { return addPeople_(b.people, who); });
    return { ok: true, added: result.added, skipped: result.skipped, forms: forms_() };
  },
  signed: function (b, who) {
    locked_(function () { signedUpload_(b, who); });
    return { ok: true, forms: forms_() };
  },
  forms: function () { return { ok: true, today: day_(new Date()), forms: forms_() }; },
  signout: function (b) {
    PropertiesService.getScriptProperties().deleteProperty('session:' + hash_(b.token));
    return { ok: true };
  },
  person: function (b) {
    return { ok: true, person: locked_(function () { return savePerson_(b); }), forms: forms_() };
  },
  prepare: function (b, who) {
    var row = Number(b.person);
    if (!Number.isInteger(row) || row <= PEOPLE_HEAD) throw said_('That person is not on the list.');
    var ids = theIds_(b.ids);
    return { ok: true, items: locked_(function () { return prepare(row, ids, who); }), forms: forms_() };
  },
  send: function (b) {
    var rows = (b.rows || []).map(theRow_);
    return { ok: true, sent: locked_(function () { return sendNow(rows); }), forms: forms_() };
  },
  signSend: function (b) {
    var row = Number(b.person);
    if (!Number.isInteger(row) || row <= PEOPLE_HEAD) throw said_('That person is not on the list.');
    var ids = theIds_(b.ids);
    return { ok: true, items: locked_(function () { return signSend_(row, ids); }), forms: forms_() };
  },
  counterForm: function (b, who) {
    var v = tab_(TABS.tracker).getRange(theRow_(b.row), 1, 1, TRACKER.length).getValues()[0];
    var id = fileId_(v[T['Signed copy']]);
    if (!id) throw said_('There is no signed copy to sign yet.');
    var bytes = DriveApp.getFileById(id).getBlob().getBytes();
    return { ok: true, title: v[T['Form']], person: v[T['Sent to']], signer: who, kinds: ['harmony', 'supervisor'],
             pdf: Utilities.base64Encode(bytes), sha: fingerprint_(bytes) };
  },
  counterDone: function (b, who) {
    var r = theRow_(b.row);
    locked_(function () { counterSigned_(r, b, who); });
    return { ok: true };
  },
  requested: function (b) {
    var r = theRow_(b.row);
    locked_(function () { markRequested(r); });
    return { ok: true, forms: forms_() };
  },
  handed: function (b, who) {
    var row = Number(b.person), ids = theIds_(b.ids);
    locked_(function () { handedOver_(row, ids, who); });
    return { ok: true, forms: forms_() };
  },
  back: function (b, who) {
    var r = theRow_(b.row);
    locked_(function () { stamp_(r, 'Returned on', b.undo ? '' : new Date(), who); });
    return { ok: true, forms: forms_() };
  },
  approve: function (b, who) {
    var r = theRow_(b.row);
    locked_(function () {
      tab_(TABS.tracker).getRange(r, T['Approved by'] + 1).setValue(b.undo ? '' : who);
      stamp_(r, 'Approved on', b.undo ? '' : new Date(), who);
    });
    return { ok: true, forms: forms_() };
  },
  reviewed: function (b, who) {
    var no = Number(b.no);
    if (!Number.isInteger(no) || no < 1) throw said_('That document is not on the list.');
    locked_(function () { log_('Reviewed by', who + ', No. ' + no + ': ' + reviewed_(no)); });
    SpreadsheetApp.flush();  // Next due and Status are formulas: read them after they recalculate
    return { ok: true, today: day_(new Date()), checklist: checklist_() };
  }
};

function stamp_(r, field, when, who) {
  var cell = tab_(TABS.tracker).getRange(r, T[field] + 1);
  if (when) date_(cell, when); else cell.setValue('');
  log_(field, 'Row ' + r + (when ? '' : ' cleared') + ' by ' + who);
}

/** A form given over in person: nothing is emailed, and it is tracked like any other. */
function handedOver_(row, ids, who) {
  var person = person_(row);
  if (!person.name) throw said_('That person is not on the list.');
  var forms = listFor_(person.role).reduce(function (all, g) { return all.concat(g.forms); }, []);
  ids.forEach(function (id) {
    var form = forms.filter(function (f) { return f.id === id; })[0];
    if (!form) throw said_('That form is not on the ' + person.role + ' list.');
    var r = trackRow_({ to: person.name, who: person.role, email: person.email, form: form.title, need: form.need,
                        copy: openLink_(id), how: BY_HAND, step: '', testTo: '' });
    date_(tab_(TABS.tracker).getRange(r, T['Date sent'] + 1), new Date());
    log_('Handed over', form.title + ' to ' + person.name + ', recorded by ' + who);
  });
}

/** One person's details, checked and made safe for a cell. Throws in words a person can act on. */
function cleanPerson_(b) {
  var out = {
    name: words_(b.name, 120), email: String(b.email || '').trim().slice(0, 254), role: String(b.role || ''),
    guardian: words_(b.guardian, 120), guardianEmail: String(b.guardianEmail || '').trim().slice(0, 254),
    start: words_(b.start, 40)
  };
  if (!out.name) throw said_('Give their name.');
  if (ROLES.indexOf(out.role) === -1) throw said_('Choose who they are.');
  if (out.email && !isEmail_(out.email)) throw said_("That email address doesn't look right.");
  if (out.guardianEmail && !isEmail_(out.guardianEmail)) throw said_("The guardian's email address doesn't look right.");
  if (!out.email && !out.guardianEmail) throw said_('Give an email address, or one for whoever signs for them.');
  return out;
}

function peopleNames_() {
  var sheet = tab_(TABS.people), last = Math.max(sheet.getLastRow(), PEOPLE_HEAD);
  return last > PEOPLE_HEAD ? sheet.getRange(PEOPLE_HEAD + 1, 1, last - PEOPLE_HEAD, 1).getValues()
    .map(function (r) { return String(r[0]).trim().toLowerCase(); }) : [];
}

function savePerson_(b) {
  var sheet = tab_(TABS.people), x = cleanPerson_(b), row = Number(b.row) || 0;
  var names = peopleNames_(), last = PEOPLE_HEAD + names.length;
  if (row) {
    if (row <= PEOPLE_HEAD || row > last) throw said_('That person is not on the list.');
    if (names[row - PEOPLE_HEAD - 1] !== x.name.toLowerCase()) {
      throw said_('A name cannot change here, because their forms are filed under it.');  // 'Sent and returned' holds names
    }
  } else {
    if (names.indexOf(x.name.toLowerCase()) !== -1) throw said_(x.name + ' is already on your list.');
    row = last + 1;
    sheet.getRange(row, P['Name'] + 1).setValue(x.name);
  }
  sheet.getRange(row, P['Email'] + 1, 1, 4).setValues([[x.email, x.role, x.guardian, x.guardianEmail]]);
  sheet.getRange(row, P['Active'] + 1).setValue(b.active === false ? 'No' : 'Yes');
  setUpPerson_(person_(row), false);
  log_('Person saved', x.name);
  return row;
}

/** A whole list at once. Each good line goes in; a bad one comes back with why. The timer makes the folders. */
function addPeople_(lines, who) {
  if (!Array.isArray(lines) || !lines.length) throw said_('There is nobody on that list.');
  if (lines.length > 500) throw said_('That is more than 500 people. Add them in smaller lists.');
  var names = peopleNames_(), good = [], skipped = [];
  lines.forEach(function (line) {
    try {
      var x = cleanPerson_(line);
      if (names.indexOf(x.name.toLowerCase()) !== -1) throw said_('Already on your list.');
      names.push(x.name.toLowerCase());
      good.push([x.name, x.email, x.role, x.guardian, x.guardianEmail, x.start, 'Yes', '', '']);
    } catch (err) {
      skipped.push({ name: String(line.name || '(no name)').slice(0, 120), why: err.message });
    }
  });
  if (good.length) {
    tab_(TABS.people).getRange(PEOPLE_HEAD + names.length - good.length + 1, 1, good.length, PEOPLE.length).setValues(good);
    log_('People added', good.length + ' by ' + who);
  }
  return { added: good.length, skipped: skipped };
}

/** A signed copy someone scanned or photographed: filed like one that came back by email. */
function signedUpload_(b, who) {
  var r = theRow_(b.row), tracker = tab_(TABS.tracker);
  var v = tracker.getRange(r, 1, 1, TRACKER.length).getValues()[0];
  var ext = (String(b.name || '').match(/\.([A-Za-z0-9]{1,5})$/) || [])[1];
  ext = ext ? ext.toLowerCase() : '';
  if (['pdf', 'jpg', 'jpeg', 'png', 'heic', 'doc', 'docx'].indexOf(ext) === -1) throw said_('Use a PDF, a photo or a Word file.');
  if (typeof b.data !== 'string' || !b.data) throw said_('That file was empty.');
  if (b.data.length > 20000000) throw said_('That file is too big. Keep it under 15 MB.');
  var name = v[T['Form']] + ' - ' + v[T['Sent to']] + ' - signed.' + ext;
  var blob = Utilities.newBlob(Utilities.base64Decode(b.data), MimeType_(ext), name);
  var file = doneFolder_(byName_(v[T['Sent to']])).createFile(blob).setDescription(MARK + 'signed copy added by ' + who);
  listCompleted_(file);
  returned_(tracker, r, v, file.getUrl(), '');
}

/* ------------------------------------------------------------------ the sign page */

// A person signs at hc.riseos.care/tool/sign from a link in their email, with no account and no
// permission screen. The link is the only way in: a long random key, kept here only as a hash, good for
// one person's own forms and nothing else. The page reads the form and stamps the signed PDF in the
// browser; this script holds no copy of a signature, only the signed PDF it files in their folder.

var SIGNING = { signOpen: signOpen_, signForm: signForm_, signDone: signDone_ };

function newLink_(name) {
  var link = (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '');
  PropertiesService.getScriptProperties().setProperty('link:' + hash_(link),
    JSON.stringify({ name: name, until: Date.now() + LINK_DAYS * 86400000 }));
  return link;
}

function linkOwner_(link) {
  var text = typeof link === 'string' && link.length === 64
    ? PropertiesService.getScriptProperties().getProperty('link:' + hash_(link)) : null;
  var held = text ? JSON.parse(text) : null;
  if (!held || held.until < Date.now()) throw said_('This link has ended. Ask Harmony Care to send you a new one.');
  return held.name;
}

/** Everything this person has been sent to sign, whichever link they came in on. */
function signRows_(name) {
  var tracker = tab_(TABS.tracker), last = Math.max(tracker.getLastRow(), TRACKER_FIRST), out = [];
  tracker.getRange(TRACKER_FIRST, 1, last - TRACKER_FIRST + 1, TRACKER.length).getValues().forEach(function (v, i) {
    if (v[T['Sent to']] === name && v[T['How it went out']] === SIGN_PAGE && v[T['Date sent']]) out.push({ r: TRACKER_FIRST + i, v: v });
  });
  return out;
}

function signRow_(body) {
  var name = linkOwner_(body.link), r = Number(body.row);
  var match = signRows_(name).filter(function (x) { return x.r === r; })[0];
  if (!match) throw said_('That form is not on your list.');
  return match;
}

function signerOf_(person) {
  var guardian = person.guardian && person.guardianEmail;
  if (person.role === 'Participant') return { name: guardian ? person.guardian : person.name, kinds: guardian ? ['representative', 'client'] : ['client', 'representative'] };
  return { name: person.name, kinds: person.role === 'Worker' ? ['worker'] : ['practitioner'] };
}

function signOpen_(body) {
  var name = linkOwner_(body.link), signer = signerOf_(byName_(name));
  return {
    ok: true, first: signer.name.split(' ')[0], signer: signer.name, kinds: signer.kinds,
    forms: signRows_(name).map(function (x) {
      return { row: x.r, title: x.v[T['Form']], need: String(x.v[T['They need to']]), done: !!x.v[T['Returned on']] };
    })
  };
}

function signForm_(body) {
  var x = signRow_(body);
  var bytes = DriveApp.getFileById(fileId_(x.v[T['Copy sent (opens it)']])).getBlob().getBytes();
  return { ok: true, title: x.v[T['Form']], pdf: Utilities.base64Encode(bytes), sha: fingerprint_(bytes) };
}

function signDone_(body) {
  return locked_(function () {
    var x = signRow_(body);
    if (x.v[T['Returned on']]) return { ok: true };  // a second press of Finish
    var signer = String(body.name || '').trim().slice(0, 120);
    if (!signer) throw said_('Your name is needed on the form.');
    var file = doneFolder_(byName_(x.v[T['Sent to']])).createFile(signedBlob_(body.pdf, x.v))
      .setDescription(MARK + 'signed on the sign page');
    listCompleted_(file);
    returned_(tab_(TABS.tracker), x.r, x.v, file.getUrl(), 'Signed on the sign page by ' + signer + '.');
    log_('Signed', x.v[T['Form']] + ', ' + x.v[T['Sent to']]);
    return { ok: true };
  });
}

function signedBlob_(data, v) {
  if (typeof data !== 'string' || !data) throw said_('The signed form did not arrive. Press Finish again.');
  if (data.length > 30000000) throw said_('That form is too big to save.');
  var bytes = Utilities.base64Decode(data);
  if (String.fromCharCode(bytes[0], bytes[1], bytes[2], bytes[3]) !== '%PDF') throw said_('The signed form did not arrive. Press Finish again.');
  return Utilities.newBlob(bytes, 'application/pdf', v[T['Form']] + ' - ' + v[T['Sent to']] + ' - signed.pdf');
}

/** Harmony's own signature goes on the signed copy, which replaces the one before it. Signing it approves it. */
function counterSigned_(r, b, who) {
  var tracker = tab_(TABS.tracker), v = tracker.getRange(r, 1, 1, TRACKER.length).getValues()[0];
  var old = DriveApp.getFileById(fileId_(v[T['Signed copy']]));
  var file = doneFolder_(byName_(v[T['Sent to']])).createFile(signedBlob_(b.pdf, v))
    .setDescription(MARK + 'signed for Harmony by ' + who);
  var shortcuts = allCompleted_().getFilesByName(old.getName());
  while (shortcuts.hasNext()) shortcuts.next().setTrashed(true);
  old.setTrashed(true);
  listCompleted_(file);
  tracker.getRange(r, T['Signed copy'] + 1).setValue(file.getUrl());
  tracker.getRange(r, T['Approved by'] + 1).setValue(who);
  stamp_(r, 'Approved on', new Date(), who);
  log_('Signed for Harmony', v[T['Form']] + ', ' + v[T['Sent to']] + ' by ' + who);
}

function fingerprint_(bytes) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, bytes).map(function (b) {
    return ('0' + (b & 255).toString(16)).slice(-2);
  }).join('');
}

/** Queues the chosen forms: what they read goes as one email of links, what they sign behind one link. */
function signSend_(row, ids) {
  var person = setUpPerson_(person_(row), false);
  var who = signers_(person);
  if (who.problems.length) throw new Error(who.problems.join(' '));
  var forms = sendGroups_(SEND_TABS[person.role]).reduce(function (all, g) { return all.concat(g.forms); }, []);
  var started = Date.now(), out = [];
  ids.forEach(function (id) {
    var form = forms.filter(function (f) { return f.id === id; })[0];
    if (!form) return out.push({ title: id, error: 'That form is not on their list.' });
    if (Date.now() - started > TIME_BUDGET_MS) return out.push({ title: form.title, error: 'Not done in time. Tick it and send it again.' });
    try {
      var reading = readOnly_(form.need);
      var r = trackRow_({
        to: person.name, who: person.role, email: who.to.join(', '), form: form.title, need: form.need,
        copy: reading ? openLink_(id) : 'https://drive.google.com/file/d/' + formPdf_(id, form.title) + '/view',
        how: reading ? EMAIL : SIGN_PAGE, step: 'Send', testTo: ''
      });
      out.push({ title: form.title, trackerRow: r });
    } catch (err) {
      log_('Not queued', form.title + ' for ' + person.name + ': ' + err.message);
      out.push({ title: form.title, error: err.message });
    }
  });
  return out;
}

/** One email and one link for everything queued for this person. */
function sendSignLink_(tracker, name, rows) {
  var mark = function (value) {
    rows.forEach(function (x) { tracker.getRange(x.r, T['Send step'] + 1).setValue(value); });
  };
  try {
    var person = byName_(name), who = signers_(person);
    if (who.problems.length) throw new Error(who.problems.join(' '));
    var mail = signEmail_(person, rows, newLink_(name));
    var thread = GmailApp.createDraft(who.to.join(','), mail.subject, mail.body, { name: setting_('Send emails as') }).send().getThread();
    thread.addLabel(GmailApp.getUserLabelByName(LABEL) || GmailApp.createLabel(LABEL));
    rows.forEach(function (x) {
      date_(tracker.getRange(x.r, T['Date sent'] + 1), new Date());
      tracker.getRange(x.r, T['Email thread'] + 1).setValue(thread.getId());
    });
    mark('Sent');
    log_('Sent', rows.length + ' to sign, to ' + who.to.join(', '));
  } catch (err) {
    mark('Not sent: ' + err.message);
    log_('Send failed', name + ', ' + rows.length + ' to sign: ' + err.message);
    tell_('Not sent: forms to sign for ' + name, err.message + '\n\n' + personLink_(name));
  }
}

function signEmail_(person, rows, link) {
  var first = signerOf_(person).name.split(' ')[0];
  return {
    subject: 'Harmony Care: ' + (rows.length === 1 ? rows[0].v[T['Form']] + ' to sign' : rows.length + ' forms to sign'),
    body: 'Hi ' + first + ',\n\nHarmony Care has ' + (rows.length === 1 ? 'a form' : rows.length + ' forms') +
          ' for you to fill in and sign. It shows you where, and you can do it on your phone or computer. ' +
          'You do not need an account.\n\nOpen ' + (rows.length === 1 ? 'it' : 'them') + ' here:\n' + SIGN_URL + link + '\n\n' +
          rows.map(function (x) { return '- ' + x.v[T['Form']]; }).join('\n') + '\n\n' +
          'Any questions, just reply here.\n\nThank you,\nHarmony Care'
  };
}

/** The PDF of a form, made once and kept, so sending takes no time and everyone signs the same audited file. */
function formPdf_(id, title) {
  var props = PropertiesService.getScriptProperties(), key = 'pdf:' + id;
  var file = DriveApp.getFileById(id), stamp = file.getLastUpdated().getTime();
  var known = JSON.parse(props.getProperty(key) || 'null');
  if (known && known.stamp === stamp) return known.pdf;
  var folder = childFolder_(DriveApp.getFolderById(requireFolder_('Signing templates folder')), 'Forms as PDFs');
  var pdf = makePdf_(file, id, title, folder);
  props.setProperty(key, JSON.stringify({ pdf: pdf, stamp: stamp }));
  return pdf;
}

function makePdf_(file, id, title, folder) {
  var type = file.getMimeType();
  if (type === 'application/pdf') return folder.createFile(file.getBlob().setName(title + '.pdf')).getId();
  if (type.indexOf('google-apps') !== -1) return folder.createFile(file.getAs('application/pdf').setName(title + '.pdf')).getId();
  var doc = DriveApp.getFileById(toDoc_(id, title + ' (making the PDF)', folder.getId()));
  var pdf = folder.createFile(doc.getAs('application/pdf').setName(title + '.pdf')).getId();
  doc.setTrashed(true);
  return pdf;
}

/** Makes the PDFs ahead of time, a few each minute, so the first send to anyone is as quick as the rest. */
function pdfsWaiting_() {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty('PDFS_READY')) return;
  var started = Date.now(), left = false;
  Object.keys(SEND_TABS).forEach(function (role) {
    sendGroups_(SEND_TABS[role]).forEach(function (g) {
      g.forms.forEach(function (f) {
        if (readOnly_(f.need) || props.getProperty('pdf:' + f.id)) return;
        if (Date.now() - started > 40000) { left = true; return; }
        try {
          formPdf_(f.id, f.title);
        } catch (err) {
          log_('No PDF yet', f.title + ': ' + err.message);
          left = true;
        }
      });
    });
  });
  if (!left) props.setProperty('PDFS_READY', '1');
}

function MimeType_(ext) {
  return { pdf: 'application/pdf', jpg: 'image/jpeg', jpeg: 'image/jpeg', png: 'image/png', heic: 'image/heic',
           doc: 'application/msword',
           docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' }[ext];
}

/** Everything the Forms screen shows, worked out from the tabs on each read so nothing stored goes stale. */
function forms_() {
  var people = [], sheet = tab_(TABS.people), last = sheet.getLastRow();
  for (var r = PEOPLE_HEAD + 1; r <= last; r++) {
    var p = person_(r);
    if (p.name) {
      people.push({ row: r, name: p.name, email: p.email, role: p.role, guardian: p.guardian,
                    guardianEmail: p.guardianEmail, active: p.active, problems: signers_(p).problems,
                    folder: p.folder ? 'https://drive.google.com/drive/folders/' + p.folder : '' });
    }
  }
  var tracker = tab_(TABS.tracker), end = Math.max(tracker.getLastRow(), TRACKER_FIRST), today = day_(new Date());
  var days = Number(setting_('Chase after (days)')) || 14, sends = [];
  tracker.getRange(TRACKER_FIRST, 1, end - TRACKER_FIRST + 1, TRACKER.length).getValues().forEach(function (v, i) {
    if (!v[T['Form']]) return;
    sends.push({
      row: TRACKER_FIRST + i, to: String(v[T['Sent to']]), form: String(v[T['Form']]), need: String(v[T['They need to']]),
      how: String(v[T['How it went out']]), standing: standing_(v, today, days), step: String(v[T['Send step']]),
      sent: v[T['Date sent']] ? day_(v[T['Date sent']]) : '', back: v[T['Returned on']] ? day_(v[T['Returned on']]) : '',
      approvedBy: String(v[T['Approved by']] || ''), approved: v[T['Approved on']] ? day_(v[T['Approved on']]) : '',
      copy: String(v[T['Copy sent (opens it)']] || ''), signed: String(v[T['Signed copy']] || ''),
      notes: String(v[T['Notes']] || '')
    });
  });
  var lists = {};
  ROLES.forEach(function (role) { lists[role] = listFor_(role); });
  return { people: people, lists: lists, sends: sends, chaseDays: days, from: setting_('Emails go out from'),
           completed: allCompleted_().getUrl() };
}

function listFor_(role) {
  var name = SEND_TABS[role];
  if (!name) throw said_('Choose who they are first.');
  var templates = templates_();
  return sendGroups_(name).map(function (g) {
    return { name: g.name, forms: g.forms.map(function (f) {
      var t = templates[f.id], reading = readOnly_(f.need);
      return { id: f.id, title: f.title, need: f.need, when: f.when, fills: f.fills, signs: f.signs, reading: reading,
               how: t && t.ready && !reading ? ESIGN : EMAIL, url: openLink_(f.id) };
    }) };
  });
}

/** Where a row stands: ready, failed, out, chase, read, back or done. */
function standing_(v, today, days) {
  if (v[T['Approved on']]) return 'done';
  if (v[T['Returned on']]) return 'back';
  if (String(v[T['Send step']]).indexOf('Not sent') === 0) return 'failed';
  if (!v[T['Date sent']]) return 'ready';
  if (readOnly_(v[T['They need to']])) return 'read';
  return daysBetween_(day_(v[T['Date sent']]), today) >= days ? 'chase' : 'out';
}

/** The Everything tab as modules, standards and items, the merged cells read back into their blocks. */
function checklist_() {
  var sheet = tab_(TABS.everything), head = 5, last = sheet.getLastRow();
  var range = sheet.getRange(head, 1, last - head + 1, 12);
  var v = range.getValues(), links = range.getRichTextValues();
  var c = index_(v[0].map(String));
  ['Practice Standard', 'Outcome', 'Indicators to be demonstrated', 'Policies and other documents',
   'Evidence/Information to be provided by Harmony Care', 'Who it is for', 'When it is needed', 'Details', 'How often',
   'Next due', 'Status'].forEach(function (h) {
    if (!(h in c)) throw new Error("The Everything tab has changed: no '" + h + "' column.");
  });
  var docs = c['Policies and other documents'], proof = c['Evidence/Information to be provided by Harmony Care'];
  var register = register_();
  var modules = [], mod = null, std = null;
  for (var i = 1; i < v.length; i++) {
    var row = v[i], a = String(row[c['Practice Standard']] || '').trim();
    var rest = row.slice(1).some(function (x) { return x !== ''; });
    if (a && !rest) {
      mod = { name: a, standards: [] };
      modules.push(mod);
      std = null;
      continue;
    }
    if (a) {
      std = { name: a, outcome: String(row[c['Outcome']] || ''), indicators: [], also: [], items: [] };
      mod.standards.push(std);
    }
    if (!std) continue;
    var indicator = String(row[c['Indicators to be demonstrated']] || '').trim();
    if (indicator) std.indicators.push(indicator);
    var col = row[docs] ? docs : row[proof] ? proof : -1;
    if (col === -1) continue;
    var link = links[i][col] && links[i][col].getLinkUrl();
    var safe = /^https:/.test(String(link || '')) ? openLink_(link) : '';
    var also = String(row[col]).match(/^Also meets this standard: (.*)$/);
    if (also) {
      std.also.push({ title: also[1], url: safe });
      continue;
    }
    var named = numberOf_(row[col]), due = row[c['Next due']];
    std.items.push({
      no: named.no, title: named.title, evidence: col === proof, url: safe,
      who: String(row[c['Who it is for']] || ''), when: String(row[c['When it is needed']] || ''),
      details: String(row[c['Details']] || ''), often: String(row[c['How often']] || ''),
      due: due ? day_(due) : '', status: String(row[c['Status']] || ''), record: register[named.no] || null
    });
  }
  return { modules: modules };
}

/** What the Checklist tab holds on each document, by its number. */
function register_() {
  var sheet = tab_(TABS.checklist), last = sheet.getLastRow(), out = {};
  var v = sheet.getRange(5, 1, last - 4, 22).getValues(), c = index_(v[0].map(String));
  var text = function (x) { return x instanceof Date ? day_(x) : String(x === '' || x == null ? '' : x); };
  for (var i = 1; i < v.length; i++) {
    if (typeof v[i][0] !== 'number') continue;
    out[v[i][0]] = {
      type: text(v[i][c['Type']]), version: text(v[i][c['Version']]), approvedBy: text(v[i][c['Approved by']]), approved: text(v[i][c['Approval date']]),
      reviewed: text(v[i][c['Last reviewed (type the date)']]), notes: text(v[i][c['Notes for the next audit']]),
      kept: text(v[i][c['Kept where']]), rule: text(v[i][c["Harmony's rule (source)"]]),
      completedBy: text(v[i][c['Completed by']]), signedBy: text(v[i][c['Signed by']])
    };
  }
  return out;
}

/** The notes and folders on Start here, for Today. */
function startNotes_() {
  var sheet = tab_('Start here'), last = sheet.getLastRow();
  var range = sheet.getRange(1, 1, last, 2), v = range.getValues(), links = range.getRichTextValues();
  var notes = [], files = [], at = '';
  v.forEach(function (row, i) {
    var a = String(row[0] || '').trim(), b = String(row[1] || '').trim();
    if (a === 'Needs attention now' || a === 'The files') { at = a; return; }
    if (!a || !b) return;
    if (at === 'Needs attention now') notes.push({ title: a, text: b });
    if (at === 'The files') files.push({ title: a, text: b, url: links[i][1] && links[i][1].getLinkUrl() });
  });
  return { attention: notes, files: files };
}

function listOf_(tabName) {
  return sendGroups_(tabName).map(function (g) {
    return { name: g.name, forms: g.forms.map(function (f) {
      return { title: f.title, need: f.need, when: f.when, fills: f.fills, signs: f.signs, url: openLink_(f.id) };
    }) };
  });
}

function planner_() {
  var sheet = tab_('Year planner'), last = sheet.getLastRow();
  var range = sheet.getRange(4, 1, last - 3, 5), v = range.getValues(), links = range.getRichTextValues();
  var jobs = [];
  v.forEach(function (row, i) {
    if (!row[1]) return;
    var due = row[0] instanceof Date || /^\d{4}-\d{2}-\d{2}/.test(String(row[0]));
    jobs.push({ row: i + 4, due: due ? day_(row[0]) : '', every: due ? '' : String(row[0]), what: String(row[1]),
                open: String(row[2] || ''), url: links[i][2] && links[i][2].getLinkUrl() ? openLink_(links[i][2].getLinkUrl()) : '', why: String(row[3] || ''),
                done: row[4] ? day_(row[4]) : '' });
  });
  return jobs;
}

function rules_() {
  var sheet = tab_('Rules and sources'), last = sheet.getLastRow();
  var range = sheet.getRange(4, 1, last - 3, 6), v = range.getValues(), links = range.getRichTextValues();
  return v.filter(function (r) { return r[0]; }).map(function (r, i) {
    return { item: String(r[0]), who: String(r[1] || ''), valid: String(r[2] || ''), keep: String(r[3] || ''),
             change: String(r[4] || ''), source: String(r[5] || ''), url: links[i][5] && links[i][5].getLinkUrl() || '' };
  });
}

function theWorker_(row) {
  var r = Number(row);
  if (!Number.isInteger(r) || r <= PEOPLE_HEAD) throw said_('That person is not on the list.');
  var person = person_(r);
  if (!person.name || person.role !== 'Worker') throw said_('That person is not a worker on the list.');
  return setUpPerson_(person, false);
}

/** Every worker, and how their file stands, read from their own tab. */
function team_() {
  var sheet = tab_(TABS.people), last = sheet.getLastRow(), out = [], today = day_(new Date());
  for (var r = PEOPLE_HEAD + 1; r <= last; r++) {
    var p = person_(r);
    if (!p.name || p.role !== 'Worker' || !p.active) continue;
    var tab = book_().getSheetByName(sheetName_(p.name)), line = { row: r, name: p.name, missing: 0, expired: 0, held: 0, next: '', nextItem: '' };
    if (tab) {
      tab.getRange(22, 2, tab.getLastRow() - 21, 8).getValues().forEach(function (v) {
        var status = String(v[7]), expiry = v[6] ? day_(v[6]) : '';
        if (status === 'Not recorded yet') line.missing++;
        else if (status === 'Expired') line.expired++;
        else if (status && status !== 'Not needed') line.held++;
        if (expiry && expiry >= today && (!line.next || expiry < line.next)) { line.next = expiry; line.nextItem = String(v[0]); }
      });
    }
    out.push(line);
  }
  return out;
}

function staffFile_(worker) {
  var tab = book_().getSheetByName(sheetName_(worker.name));
  if (!tab) throw said_(worker.name + "'s staff file is still being made. Try again in a minute.");
  var flags = tab.getRange(9, 1, 11, 2).getValues().map(function (v, i) { return { row: 9 + i, label: String(v[0]), yes: v[1] === 'Yes' }; });
  var range = tab.getRange(22, 1, tab.getLastRow() - 21, 11), v = range.getValues(), f = range.getFormulas(), links = range.getRichTextValues();
  var items = [];
  v.forEach(function (row, i) {
    if (!row[1]) return;
    items.push({ row: 22 + i, group: String(row[0]), item: String(row[1]), keep: String(row[3] || ''), form: String(row[4] || ''),
                 formUrl: links[i][4] && links[i][4].getLinkUrl() ? openLink_(links[i][4].getLinkUrl()) : '', renew: String(row[5] || ''), done: row[6] ? day_(row[6]) : '',
                 expiry: row[7] ? day_(row[7]) : '', expiryWorkedOut: !!f[i][7], status: String(row[8] || '') });
  });
  return { name: worker.name, flags: flags, items: items, folder: 'https://drive.google.com/drive/folders/' + worker.folder };
}

function staffDate_(worker, b, who) {
  var tab = book_().getSheetByName(sheetName_(worker.name)), row = Number(b.sheetRow), value = String(b.value || '');
  if (!tab || !Number.isInteger(row) || row < 22 || row > tab.getLastRow() || !tab.getRange(row, 2).getValue()) throw said_('That is not on their file.');
  if (value && !/^\d{4}-\d{2}-\d{2}$/.test(value)) throw said_("That date doesn't look right.");
  var col = b.field === 'expiry' ? 8 : 7, cell = tab.getRange(row, col);
  if (col === 8 && tab.getRange(row, 8).getFormula()) throw said_('That expiry works itself out from the date done.');
  if (value) cell.setValue(value).setNumberFormat('d mmm yyyy'); else cell.setValue('');
  log_('Staff file', worker.name + ', row ' + row + ' by ' + who);
}
