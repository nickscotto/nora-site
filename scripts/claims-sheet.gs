/**
 * Google Apps Script for the "Nora - $100 claims" sheet.
 * The site POSTs each claim here (see the claim form handler in index.html);
 * this adds it as a new row at the top with an unchecked "Texted" box.
 *
 * Setup: Extensions > Apps Script, paste this file, run setup() once,
 * then Deploy > New deployment > Web app (Execute as: Me, Who has access: Anyone).
 */

const SHEET_NAME = "Claims";
const HEADERS = ["Claimed", "Name", "Phone", "Text her", "Texted", "Notes"];

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME, 0);
  sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS])
    .setFontWeight("bold").setBackground("#4A2F22").setFontColor("#F3ECE3");
  sh.setFrozenRows(1);
  sh.setColumnWidth(1, 150); sh.setColumnWidth(2, 120); sh.setColumnWidth(3, 130);
  sh.setColumnWidth(4, 90); sh.setColumnWidth(5, 70); sh.setColumnWidth(6, 260);
  // Gray out rows once "Texted" is checked
  const rule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied("=$E2=TRUE").setFontColor("#9A8B80")
    .setRanges([sh.getRange("A2:F1000")]).build();
  sh.setConditionalFormatRules([rule]);
  const def = ss.getSheetByName("Sheet1");
  if (def && def.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(def);
}

function doPost(e) {
  let d = {};
  try { d = JSON.parse(e.postData.contents); } catch (err) { return out("bad json"); }
  if (d._honey) return out("ok"); // bot filled the hidden field
  const name = clean(d.first_name).slice(0, 40);
  const phone = String(d.phone || "").replace(/\D/g, "").replace(/^1(?=\d{10}$)/, "");
  if (!name || phone.length !== 10) return out("invalid");

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    const pretty = "(" + phone.slice(0, 3) + ") " + phone.slice(3, 6) + "-" + phone.slice(6);
    sh.insertRowAfter(1);
    sh.getRange(2, 1, 1, 6).setValues([[
      new Date(), name, pretty,
      '=HYPERLINK("sms:+1' + phone + '","Text")', false, ""
    ]]).setFontWeight("normal").setBackground(null).setFontColor(null);
    sh.getRange(2, 1).setNumberFormat("ddd m/d h:mm am/pm");
    sh.getRange(2, 5).insertCheckboxes();
  } finally {
    lock.releaseLock();
  }
  return out("ok");
}

// Stop names like "=IMPORTXML(...)" from running as formulas
function clean(v) {
  const s = String(v || "").trim();
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function out(msg) {
  return ContentService.createTextOutput(JSON.stringify({ result: msg })).setMimeType(ContentService.MimeType.JSON);
}
