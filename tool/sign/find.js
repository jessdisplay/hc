/* Reads a form's PDF text and works out what a person does on it: where to tick, write and sign.
   Harmony's forms come from many templates, so nothing here is set up per form. It takes pdf.js text items
   and returns spots in PDF points (origin bottom left). No page or network here, so it runs in the tests. */
(function (root) {
  'use strict';
  var BOX = /[☐☑☒□▢⬜]/;
  var RUN = /(?:[….·]{2,}\/){2}[….·]{2,}|[….·]{6,}|…{3,}|_{5,}/g;
  var DATED = /\//;
  var ROLES = [
    ['worker', /client.?s worker/i],
    ['', /interpreter/i],
    ['representative', /signed by the representative|executed by the representative|advocates?\s*\/\s*guardian|supporter|guardian.?s signature|parent.*consent/i],
    ['harmony', /on behalf|services\)?,? by:?$|by:$|organisation representative|authorised harmony|executed by silp|key management|service provider/i],
    ['supervisor', /supervisor|appraiser|manager|principal|assessor/i],
    ['client', /signed by the client|executed by the client|^client|client.?s signature|client \/ client/i],
    ['worker', /signed by the worker|worker|staff member|staff attending|employee|appraisee|applicant|tester/i],
    ['practitioner', /practitioner|prescriber|developer|trainer|health professional/i]
  ];

  function itemsOf(textItems) {
    return textItems.filter(function (i) { return i.str && i.str.trim(); }).map(function (i) {
      return { s: i.str, x: i.transform[4], y: i.transform[5], w: i.width, h: Math.abs(i.transform[3]) || i.height || 10 };
    });
  }

  function linesOf(items) {
    var sorted = items.slice().sort(function (a, b) { return b.y - a.y || a.x - b.x; });
    var lines = [];
    sorted.forEach(function (it) {
      var line = lines[lines.length - 1];
      if (line && Math.abs(line.y - it.y) < 3) line.items.push(it);
      else lines.push({ y: it.y, items: [it] });
    });
    lines.forEach(function (l) {
      l.items.sort(function (a, b) { return a.x - b.x; });
      l.text = l.items.map(function (i) { return i.s; }).join(' ').replace(/\s+/g, ' ').trim();
      l.left = l.items[0].x;
    });
    return lines;
  }

  /** A role, '' for a named party that is nobody's job here, or null when the words say nothing. */
  function role(text) {
    for (var i = 0; i < ROLES.length; i++) if (ROLES[i][1].test(text)) return ROLES[i][0];
    return null;
  }

  function blanksOn(line) {
    var out = [];
    line.items.forEach(function (it) {
      var m, re = new RegExp(RUN.source, 'g');
      while ((m = re.exec(it.s))) {
        var share = it.w / it.s.length;
        var blank = { x: it.x + m.index * share, w: m[0].length * share, y: it.y, h: it.h, dated: DATED.test(m[0]) };
        if (blank.w >= 24) out.push(blank);
      }
    });
    return out;
  }

  /** The role of a signature spot: its own label first, then the lines just under it, then the heading above. */
  function roleFor(lines, index) {
    var own = role(lines[index].text);
    if (own !== null) return own;
    var i;
    for (i = index + 1; i <= index + 2 && i < lines.length; i++) {
      if (lines[i - 1].y - lines[i].y > 34) break;
      var below = lines[i].text;
      if (below.length < 70 && !/please print/i.test(below)) { var r = role(below); if (r !== null) return r; }
    }
    for (i = index - 1; i >= 0 && i >= index - 8; i--) {
      if (lines[i].text.length < 90) { var up = role(lines[i].text); if (up !== null) return up; }
    }
    return '';
  }

  /** @param {Array} textItems pdf.js items for one page @param {number} width page width in points */
  function findSpots(textItems, width) {
    var items = itemsOf(textItems), lines = linesOf(items), spots = [];
    var right = (width || 595) - 36;
    var signed = {};

    lines.forEach(function (line) {
      line.items.forEach(function (it) {
        for (var i = 0; i < it.s.length; i++) {
          if (!BOX.test(it.s.charAt(i))) continue;
          var share = it.w / it.s.length;
          var size = Math.min(it.h * 1.1, 13);
          spots.push({ type: 'tick', x: it.x + i * share, y: it.y - 1, w: size, h: size });
        }
      });
    });

    lines.forEach(function (line, index) {
      blanksOn(line).forEach(function (b) {
        var before = line.items.filter(function (i) { return i.x + 2 < b.x; }).map(function (i) { return i.s.replace(/[…._]/g, ''); }).join(' ');
        var next = lines[index + 1];
        var near = next && line.y - next.y < 26 && line.y - next.y > 3;
        var under = near ? next.text : '';
        var box = { x: b.x, y: b.y - 1, w: b.w, h: Math.max(b.h + 4, 14) };
        if (b.dated || /date:?\s*$/i.test(before)) { box.type = 'date'; spots.push(box); return; }
        if (/^signature\b/i.test(under) || /signature:?\s*$/i.test(before)) {
          box.type = 'sign';
          box.who = roleFor(lines, near && /^signature\b/i.test(under) ? index + 1 : index);
          box.h = Math.max(box.h, 26);
          signed[index] = signed[index + 1] = true;
          spots.push(box);
          return;
        }
        if (/^name\b/i.test(under) || /name:?\s*$/i.test(before)) { box.type = 'name'; spots.push(box); return; }
        box.type = 'text';
        spots.push(box);
      });
    });

    // a signature with no line to write on: a table cell beside or under its label, or a line lost to a page break
    lines.forEach(function (line, index) {
      if (signed[index]) return;
      line.items.forEach(function (it, n) {
        var s = it.s.trim();
        if (!/signature/i.test(s) || s.length > 44 || /[.]$/.test(s) || /^section/i.test(s)) return;
        var after = line.items[n + 1];
        var room = (after ? after.x : right) - (it.x + it.w);
        var spot = { type: 'sign', who: roleFor(lines, index), h: 26 };
        if (room > 90) { spot.x = it.x + it.w + 8; spot.w = Math.min(room - 12, 220); spot.y = it.y - 6; }
        else { spot.x = it.x; spot.w = Math.max(it.w, 110); spot.y = it.y - 30; }
        spots.push(spot);
      });
    });

    // 'Date:' with nothing after it, usually beside a signature
    lines.forEach(function (line) {
      var last = line.items[line.items.length - 1];
      if (!/^date:$/i.test(last.s.trim()) || right - last.x - last.w < 40) return;
      spots.push({ type: 'date', x: last.x + last.w + 4, y: last.y - 1, w: Math.min(right - last.x - last.w, 110), h: 14 });
    });
    return spots;
  }

  /** The spots this signer signs: the first of their roles the form has, or those it names no one for. */
  function mine(spots, kinds) {
    var signs = spots.filter(function (s) { return s.type === 'sign'; });
    for (var i = 0; i < kinds.length; i++) {
      var own = signs.filter(function (s) { return s.who === kinds[i]; });
      if (own.length) return own;
    }
    return signs.filter(function (s) { return !s.who; });
  }

  var api = { findSpots: findSpots, mine: mine };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.HCFind = api;
})(this);
