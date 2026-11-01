/* Reads a list of people from a spreadsheet, in the browser: nothing leaves the page until it is checked
   and someone presses Add. Excel is unzipped here, so no library comes along for one job. */
(function (root) {
  'use strict';

  function parseText(text) {
    var clean = String(text || '').replace(/^﻿/, '');
    var first = clean.split(/\r?\n/)[0] || '';
    var sep = first.indexOf('\t') !== -1 ? '\t' : ',';
    var rows = [], row = [], cell = '', quoted = false;
    for (var i = 0; i < clean.length; i++) {
      var c = clean[i];
      if (quoted) {
        if (c === '"' && clean[i + 1] === '"') { cell += '"'; i++; }
        else if (c === '"') quoted = false;
        else cell += c;
      } else if (c === '"' && !cell) quoted = true;
      else if (c === sep) { row.push(cell); cell = ''; }
      else if (c === '\n') { row.push(cell); rows.push(row); row = []; cell = ''; }
      else if (c !== '\r') cell += c;
    }
    if (cell || row.length) { row.push(cell); rows.push(row); }
    return rows.filter(function (r) { return r.some(function (x) { return String(x).trim(); }); });
  }

  /** Every file in the zip, as text, by name. */
  function unzip(buffer) {
    var view = new DataView(buffer), end = -1;
    for (var i = buffer.byteLength - 22; i >= 0 && i > buffer.byteLength - 66000; i--) {
      if (view.getUint32(i, true) === 0x06054b50) { end = i; break; }
    }
    if (end === -1) return Promise.reject(new Error('not a zip'));
    var count = view.getUint16(end + 10, true), at = view.getUint32(end + 16, true), jobs = [], names = {};
    for (var n = 0; n < count; n++) {
      var method = view.getUint16(at + 10, true), size = view.getUint32(at + 20, true);
      var nameLength = view.getUint16(at + 28, true), extra = view.getUint16(at + 30, true), comment = view.getUint16(at + 32, true);
      var local = view.getUint32(at + 42, true);
      var name = new TextDecoder().decode(new Uint8Array(buffer, at + 46, nameLength));
      var start = local + 30 + view.getUint16(local + 26, true) + view.getUint16(local + 28, true);
      jobs.push(inflate(new Uint8Array(buffer, start, size), method).then(function (nm, text) { names[nm] = text; }.bind(null, name)));
      at += 46 + nameLength + extra + comment;
    }
    return Promise.all(jobs).then(function () { return names; });
  }
  function inflate(bytes, method) {
    if (method === 0) return Promise.resolve(new TextDecoder().decode(bytes));
    var stream = new Blob([bytes]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
    return new Response(stream).text();
  }
  function column(ref) {
    var letters = String(ref).replace(/\d/g, ''), n = 0;
    for (var i = 0; i < letters.length; i++) n = n * 26 + letters.charCodeAt(i) - 64;
    return n - 1;
  }
  function readXlsx(buffer) {
    return unzip(buffer).then(function (files) {
      var parse = function (x) { return new DOMParser().parseFromString(x, 'application/xml'); };
      var shared = files['xl/sharedStrings.xml'] ? [].map.call(parse(files['xl/sharedStrings.xml']).getElementsByTagName('si'), function (si) {
        return [].map.call(si.getElementsByTagName('t'), function (t) { return t.textContent; }).join('');
      }) : [];
      var sheet = files['xl/worksheets/sheet1.xml'];
      if (!sheet) throw new Error('no sheet');
      var rows = [].map.call(parse(sheet).getElementsByTagName('row'), function (row) {
        var out = [];
        [].forEach.call(row.getElementsByTagName('c'), function (c) {
          var v = c.getElementsByTagName('v')[0], type = c.getAttribute('t'), text = '';
          if (type === 's' && v) text = shared[Number(v.textContent)] || '';
          else if (type === 'inlineStr') text = [].map.call(c.getElementsByTagName('t'), function (t) { return t.textContent; }).join('');
          else if (v) text = v.textContent;
          out[column(c.getAttribute('r'))] = text;
        });
        for (var i = 0; i < out.length; i++) if (out[i] === undefined) out[i] = '';
        return out;
      });
      return rows.filter(function (r) { return r.some(function (x) { return String(x).trim(); }); });
    });
  }

  function roleOf(word) {
    var w = String(word || '').toLowerCase();
    if (/participant|client/.test(w)) return 'Participant';
    if (/practitioner|partner|nurse|doctor|\bgp\b|therap|dietit|patholog/.test(w)) return 'Practitioner or partner';
    if (/worker|staff|employee|coordinator|manager|admin/.test(w)) return 'Worker';
    return '';
  }
  function findColumns(head) {
    var at = {};
    head.forEach(function (raw, i) {
      var h = String(raw).trim().toLowerCase();
      var signer = /guardian|representative|signs for|nominee|carer|their email/.test(h);
      if (signer && /mail/.test(h)) at.guardianEmail = i;
      else if (signer) at.guardian = i;
      else if (/mail/.test(h)) at.email = i;
      else if (/^(first|given)/.test(h)) at.first = i;
      else if (/^(last|surname|family)/.test(h)) at.last = i;
      else if (/name/.test(h)) at.name = i;
      else if (/role|type|who they are|position|group/.test(h)) at.role = i;
      else if (/start/.test(h)) at.start = i;
    });
    return at;
  }
  /** People from rows, each with what is wrong in plain words, or nothing. The script checks again. */
  function toPeople(rows, fallbackRole, known) {
    if (!rows.length) return [];
    var at = findColumns(rows[0]);
    var hasHead = at.name !== undefined || at.first !== undefined || at.email !== undefined;
    if (!hasHead) at = { name: 0, email: 1 };
    var seen = {};
    (known || []).forEach(function (n) { seen[String(n).trim().toLowerCase()] = 'list'; });
    var pick = function (r, k) { return at[k] === undefined ? '' : String(r[at[k]] === undefined ? '' : r[at[k]]).trim(); };
    return rows.slice(hasHead ? 1 : 0).map(function (r) {
      var name = pick(r, 'name') || [pick(r, 'first'), pick(r, 'last')].filter(Boolean).join(' ');
      var p = { name: name, email: pick(r, 'email'), role: roleOf(pick(r, 'role')) || fallbackRole || '', guardian: pick(r, 'guardian'),
                guardianEmail: pick(r, 'guardianEmail'), start: pick(r, 'start') };
      var good = /^[^\s@,;]+@[^\s@,;]+\.[^\s@,;]+$/, key = name.toLowerCase();
      if (!name) p.problem = 'No name';
      else if (seen[key] === 'list') p.problem = 'Already on your list';
      else if (seen[key]) p.problem = 'Twice in this file';
      else if (!p.role) p.problem = 'Say who they are';
      else if (p.email && !good.test(p.email)) p.problem = 'Email looks wrong';
      else if (p.guardianEmail && !good.test(p.guardianEmail)) p.problem = "Guardian's email looks wrong";
      else if (!p.email && !p.guardianEmail) p.problem = 'No email';
      if (name && !seen[key]) seen[key] = 'file';
      return p;
    });
  }
  root.HCPeople = { parseText: parseText, readXlsx: readXlsx, toPeople: toPeople };
})(typeof window !== 'undefined' ? window : module.exports);
