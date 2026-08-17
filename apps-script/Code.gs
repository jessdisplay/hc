/**
 * The audit room :: intake hook
 *
 * Receives a file from the website's drop panel, files it in the right place
 * inside the **Rise** folder, and appends a row to the Evidence Index so the
 * auditor gets a note and a direct link.
 *
 * Deploy: Extensions > Apps Script, paste this in, then Deploy > New deployment
 * > Web app, "Execute as: Me", "Who has access: Anyone". Copy the /exec URL.
 *
 * Before deploying, set the PIN: Project Settings > Script properties >
 * Add script property, name INTAKE_PIN, value whatever you choose. The site is
 * public, so without a PIN anyone who found the URL could post files into the
 * audit folder. The PIN is checked here on the server and never appears in the
 * page source.
 */

var RISE_FOLDER  = '1Ic6YEEUzOzIiKGOeWccdhCL0oCJaMSTk';  // **Rise**
var INTAKE       = '1vt3djHYlauK0l0wiPjBHX5yMzq57Lb8T';  // 00 Intake
var EVIDENCE     = '1wXqLJvn3EPOe-jQRm_LvxNzcpPgLs9Sl';  // 01 Evidence
var INDEX_SHEET  = '1ZgoiX8scx2vLdZbSyBSjJtGb-4JjNl3pksI5qgVfCKs';

var MAX_BYTES = 40 * 1024 * 1024;

/**
 * Who hears about a new file. Set a script property NOTIFY to a comma separated
 * list of addresses, e.g. "victoria@harmonycaresupportservices.com.au,dev@riseos.care".
 * Leave it unset and nothing is sent.
 */
function notify_(subject, body) {
  var to = PropertiesService.getScriptProperties().getProperty('NOTIFY');
  if (!to) return;
  try {
    MailApp.sendEmail({ to: to, subject: subject, body: body, name: 'The audit room' });
  } catch (err) {
    // a failed notification must never lose the file that triggered it
  }
}

/** Scope areas the panel offers, mapped to their folder ids. */
var AREAS = {
  'Core 1 - Rights and Responsibilities':            '17liW37_WtC4JY0Fxnx8iby5w01rYot5O',
  'Core 2 - Governance and Operational Management':  '1h33rR9cofnWIreEqUlqXCireV1S0jlxO',
  'Core 3 - Provision of Supports':                  '1ZTEfDRAIaxO2gWUFPnKoVeCkyz7t2h_j',
  'Core 4 - Support Provision Environment':          '1EFm2IO7urIxT_jzf2XIGaia09RqntnK0',
  'Module 1 - High Intensity Daily Personal Activities': '11ovxYlUSZX-E-W4dGbMdcJFagTwkpVBK',
  'Module 2A - Implementing Behaviour Support Plans':    '19oK8AUHrVpIownwXReIvHeKDYKXKLhHs',
  'Module 4 - Specialised Support Coordination':         '1sgVfY4GL7stq06IPfnn_1i-jVc23SKqY',
  'Module 5A - Supported Independent Living':            '1o4fos7l36OO4RuMtCim2m-cqRiP0D_j4',
  'Workforce':                                       '1T6ibl15HXwuQZm5xdEfhBubUXNaTv332',
  'Insurances and Registrations':                    '1GiNh3GbMU0P6uN3rCx7adf2XaisPFP66',
  'Registers':                                       '1Bww1Tc2jytkyaNYqKxl76Ft8nOB0acEj'
};

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/** The panel calls this on load to fill the scope dropdown. */
function doGet() {
  return json_({ ok: true, areas: Object.keys(AREAS) });
}

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);

    var pin = PropertiesService.getScriptProperties().getProperty('INTAKE_PIN');
    if (!pin) return json_({ ok: false, error: 'The intake PIN has not been set on the script yet.' });
    if (String(body.pin || '') !== pin) return json_({ ok: false, error: 'That PIN is not right.' });

    if (!body.filename) return json_({ ok: false, error: 'No file name came through.' });
    if (!body.data)     return json_({ ok: false, error: 'No file came through.' });

    var bytes = Utilities.base64Decode(body.data);
    if (bytes.length > MAX_BYTES) {
      return json_({ ok: false, error: 'That file is over 40 MB. Put it straight in the Intake folder instead.' });
    }

    // Unfiled uploads wait in Intake rather than guessing a home for themselves.
    var area   = body.area && AREAS[body.area] ? body.area : '';
    var target = area ? DriveApp.getFolderById(AREAS[area]) : DriveApp.getFolderById(INTAKE);

    var blob = Utilities.newBlob(bytes, body.mimeType || 'application/octet-stream', body.filename);
    var file = target.createFile(blob);

    var note = String(body.note || '').trim();
    if (note) file.setDescription(note);

    SpreadsheetApp.openById(INDEX_SHEET).getSheets()[0].appendRow([
      Utilities.formatDate(new Date(), 'Australia/Brisbane', 'yyyy-MM-dd'),
      body.filename,
      note,
      area || 'Not filed yet, sitting in Intake',
      area || '00 Intake',
      String(body.by || '').trim(),
      file.getUrl()
    ]);

    notify_(
      'New audit file: ' + file.getName(),
      file.getName() + '\n\n' +
      'Filed in: ' + (area || 'Intake, not filed yet') + '\n' +
      'Provided by: ' + (String(body.by || '').trim() || 'not said') + '\n' +
      'Note: ' + (note || 'none') + '\n\n' +
      'Open it: ' + file.getUrl() + '\n' +
      'The index: https://docs.google.com/spreadsheets/d/' + INDEX_SHEET + '/edit'
    );

    return json_({ ok: true, name: file.getName(), url: file.getUrl(), area: area || 'Intake' });

  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

/**
 * Catches the other half of the two-way intake: files dragged straight into the
 * Intake folder in Drive, which never touch doPost.
 *
 * Add a time-driven trigger for this (Triggers > Add trigger > watchIntake >
 * Time-driven > Minutes timer > every 5 minutes). It indexes anything new and
 * sends the same notification, so both routes behave identically.
 */
function watchIntake() {
  var props = PropertiesService.getScriptProperties();
  var seen = JSON.parse(props.getProperty('SEEN_INTAKE') || '[]');
  var sheet = SpreadsheetApp.openById(INDEX_SHEET).getSheets()[0];

  var files = DriveApp.getFolderById(INTAKE).getFiles();
  var added = [];

  while (files.hasNext()) {
    var file = files.next();
    var id = file.getId();
    if (seen.indexOf(id) !== -1) continue;

    seen.push(id);
    added.push(file);

    sheet.appendRow([
      Utilities.formatDate(new Date(), 'Australia/Brisbane', 'yyyy-MM-dd'),
      file.getName(),
      file.getDescription() || '',
      'Dropped in Drive, not filed yet',
      '00 Intake',
      '',
      file.getUrl()
    ]);
  }

  // keep the memory from growing without bound across a long audit
  props.setProperty('SEEN_INTAKE', JSON.stringify(seen.slice(-500)));

  if (added.length) {
    notify_(
      added.length + ' new file' + (added.length > 1 ? 's' : '') + ' in Intake',
      added.map(function (f) { return f.getName() + '\n' + f.getUrl(); }).join('\n\n') +
      '\n\nThe index: https://docs.google.com/spreadsheets/d/' + INDEX_SHEET + '/edit'
    );
  }
}
