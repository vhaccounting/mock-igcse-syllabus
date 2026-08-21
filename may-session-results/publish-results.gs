/**
 * MAY SESSION RESULT SHEET — build public, publishable result tabs.
 *
 * The working tabs (OL / AS / A2) hold student and parent phone numbers.
 * A page that reads a tab directly exposes EVERY column in it, because the
 * browser fetches the whole tab and anyone can read the response. So the
 * public page never points at a working tab. This script copies out only the
 * four result columns into *_PUBLIC tabs, and only those get link-shared.
 *
 * Setup: Extensions → Apps Script → paste → run buildPublicResults once,
 * then use the "Results" menu that appears on the sheet.
 */

// Only these columns are ever copied. Anything unrecognised is dropped, so a
// new column added to a working tab cannot leak by default.
var COLUMN_RULES = [
  { out: 'SL',    match: ['sl.', 'sl no', 'sl', 'serial', 'roll'] },
  { out: 'Name',  match: ['full name', 'student name', 'name'] },
  { out: 'Marks', match: ['marks', 'result', 'as ums', 'a2 ums', 'ums'] },
  { out: 'Grade', match: ['grade', 'as grade', 'a2 grade'] }
];

// Never copy a column whose header looks like a contact detail, even if some
// future rule above would otherwise match it.
var BLOCKED = /contact|phone|mobile|whatsapp|number|father|mother|parent|guardian|email|address|nid|passport/i;

var TABS = [
  { source: 'OL', target: 'OL_PUBLIC' },
  { source: 'AS', target: 'AS_PUBLIC' },
  { source: 'A2', target: 'A2_PUBLIC' }
];

function buildPublicResults() {
  var ss = SpreadsheetApp.getActive();
  var built = [];

  TABS.forEach(function (tab) {
    var source = ss.getSheetByName(tab.source);
    if (!source) return;

    var values = source.getDataRange().getDisplayValues();
    if (values.length < 2) return;

    var headers = values[0].map(function (h) { return String(h).trim().toLowerCase(); });
    var picked = [];

    COLUMN_RULES.forEach(function (rule) {
      var index = -1;
      for (var i = 0; i < rule.match.length && index === -1; i++) {
        for (var c = 0; c < headers.length; c++) {
          if (headers[c] === rule.match[i] && !BLOCKED.test(headers[c])) { index = c; break; }
        }
      }
      // The OL tab numbers students in an unlabelled column A.
      if (index === -1 && rule.out === 'SL' && !headers[0]) index = 0;
      if (index !== -1) picked.push({ out: rule.out, index: index });
    });

    var nameCol = picked.filter(function (p) { return p.out === 'Name'; })[0];
    if (!nameCol) throw new Error('No name column found on tab "' + tab.source + '"');

    var rows = [];
    for (var r = 1; r < values.length; r++) {
      var name = String(values[r][nameCol.index] || '').trim();
      if (!name || /^(grade|max mark|min raw mark)$/i.test(name)) continue;
      rows.push(picked.map(function (p) { return values[r][p.index]; }));
    }

    var target = ss.getSheetByName(tab.target) || ss.insertSheet(tab.target);
    target.clear();
    target.getRange(1, 1, 1, picked.length)
      .setValues([picked.map(function (p) { return p.out; })])
      .setFontWeight('bold').setBackground('#353434').setFontColor('#ffffff');
    if (rows.length) target.getRange(2, 1, rows.length, picked.length).setValues(rows);
    target.setFrozenRows(1);
    target.autoResizeColumns(1, picked.length);

    built.push(tab.target + ' (' + rows.length + ')');
  });

  SpreadsheetApp.getActive().toast('Rebuilt: ' + built.join(', '), 'Results', 8);
}

/** Confirms no *_PUBLIC tab carries a contact column. Run before sharing. */
function auditPublicTabs() {
  var ss = SpreadsheetApp.getActive();
  var problems = [];

  TABS.forEach(function (tab) {
    var sheet = ss.getSheetByName(tab.target);
    if (!sheet) return;
    var values = sheet.getDataRange().getDisplayValues();
    if (!values.length) return;

    values[0].forEach(function (header, c) {
      if (BLOCKED.test(String(header))) problems.push(tab.target + ': header "' + header + '"');
      for (var r = 1; r < values.length; r++) {
        var cell = String(values[r][c] || '').replace(/\D/g, '');
        if (cell.length >= 9 && /^(88)?01\d{9}$/.test(cell)) {
          problems.push(tab.target + ': row ' + (r + 1) + ' column ' + (c + 1) + ' looks like a phone number');
          break;
        }
      }
    });
  });

  SpreadsheetApp.getUi().alert(
    problems.length ? 'Do not share yet:\n\n' + problems.join('\n') : 'Clean — no contact data in the public tabs.'
  );
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Results')
    .addItem('Rebuild public result tabs', 'buildPublicResults')
    .addItem('Audit public tabs for contact data', 'auditPublicTabs')
    .addToUi();
}
