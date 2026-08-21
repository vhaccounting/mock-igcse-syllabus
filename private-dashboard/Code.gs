/**
 * Student Register — private dashboard.
 *
 * Serves the teacher dashboard (names, marks, grades, WhatsApp links to
 * parents) to a short list of signed-in Google accounts. Nothing is public.
 *
 * The browser never contacts Google Sheets. This script reads the sheet with
 * its own credentials and hands rows to the page, which is what allows the
 * spreadsheet's own sharing to be set back to Restricted.
 *
 * SETUP
 *   1. Extensions -> Apps Script, inside the marksheet spreadsheet.
 *   2. Add this file as Code.gs and the dashboard as Index.html.
 *   3. Put your own Google accounts in ALLOWLIST below.
 *   4. Deploy -> New deployment -> Web app
 *        Execute as:     Me
 *        Who has access: Anyone with a Google account
 *      "Anyone with a Google account" only means Google does the sign-in.
 *      ALLOWLIST below decides who actually gets data.
 *   5. Set the spreadsheet back to Restricted in Share.
 */

// Only these accounts may load the dashboard. Lowercase.
var ALLOWLIST = [
  'actbytak@gmail.com'
  // , 'second.teacher@gmail.com'
];

var TAB_NAMES = ['OL', 'AS', 'A2'];

function doGet() {
  var viewer = currentUser();

  if (!isAllowed(viewer)) {
    return HtmlService.createHtmlOutput(denialPage(viewer))
      .setTitle('Student Register — access required');
  }

  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Student Register — Vertical Horizon')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

function currentUser() {
  // getActiveUser() is populated for accounts in the same domain as the owner,
  // and for consumer accounts once the user has authorised the app.
  return String(Session.getActiveUser().getEmail() || '').trim().toLowerCase();
}

function isAllowed(email) {
  if (!email) return false;
  for (var i = 0; i < ALLOWLIST.length; i++) {
    if (String(ALLOWLIST[i]).trim().toLowerCase() === email) return true;
  }
  return false;
}

/**
 * Called from the page via google.script.run. Returns the tab as rows of
 * display strings, the same shape the dashboard already parses.
 */
function getRows(tabName) {
  var viewer = currentUser();
  if (!isAllowed(viewer)) throw new Error('Your account (' + (viewer || 'unknown') + ') is not on the access list.');

  // Never let the page name an arbitrary tab.
  if (TAB_NAMES.indexOf(tabName) === -1) throw new Error('Unknown tab "' + tabName + '".');

  var sheet = SpreadsheetApp.getActive().getSheetByName(tabName);
  if (!sheet) throw new Error('No tab named "' + tabName + '" in this spreadsheet.');

  logAccess(viewer, tabName);
  return sheet.getDataRange().getDisplayValues();
}

/** Append-only record of who opened what, so access is reviewable. */
function logAccess(email, tabName) {
  try {
    var ss = SpreadsheetApp.getActive();
    var log = ss.getSheetByName('AccessLog');
    if (!log) {
      log = ss.insertSheet('AccessLog');
      log.appendRow(['When', 'Account', 'Tab']);
      log.setFrozenRows(1);
    }
    log.appendRow([new Date(), email, tabName]);
  } catch (err) {
    // Logging must never block a legitimate read.
  }
}

function denialPage(viewer) {
  var who = viewer ? 'Signed in as <b>' + escapeHtml(viewer) + '</b>.' : 'You are not signed in to Google.';
  return '<!DOCTYPE html><meta name="viewport" content="width=device-width,initial-scale=1">' +
    '<div style="font:15px/1.6 system-ui,sans-serif;max-width:460px;margin:16vh auto;padding:0 24px;color:#353434">' +
    '<div style="height:4px;width:44px;background:#ee4345;border-radius:2px;margin-bottom:14px"></div>' +
    '<h1 style="font-size:21px;margin:0 0 8px">Access required</h1>' +
    '<p style="color:#5B6472">' + who + ' This dashboard is limited to Vertical Horizon staff accounts.</p>' +
    '<p style="color:#5B6472">If that should include you, ask for your address to be added to the access list.</p>' +
    '</div>';
}

function escapeHtml(s) {
  return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
  });
}
