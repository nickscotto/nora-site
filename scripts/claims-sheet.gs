/**
 * Google Apps Script for the "Nora - $100 claims" sheet.
 * The site POSTs each claim here (see the claim form handler in index.html);
 * this adds it as a new row at the top with an unchecked "Texted" box.
 * A phone number already in the sheet is not added twice; H1 shows the count of distinct people.
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
  const note = clean(d.note).slice(0, 200);
  const when = d.claimed_at ? new Date(d.claimed_at) : new Date();

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    // Count of distinct people (by phone), kept in H1 so new rows never move it
    if (!sh.getRange("H1").getFormula()) sh.getRange("H1").setFormula('="People: "&(COUNTUNIQUE(C:C)-1)').setFontWeight("bold");
    const pretty = "(" + phone.slice(0, 3) + ") " + phone.slice(3, 6) + "-" + phone.slice(6);
    // Same phone already listed: add the note to that row instead of a duplicate
    const last = sh.getLastRow();
    if (last > 1) {
      const phones = sh.getRange(2, 3, last - 1, 1).getDisplayValues();
      for (let i = 0; i < phones.length; i++) {
        if (phones[i][0] === pretty) {
          if (note) {
            const c = sh.getRange(i + 2, 6), cur = c.getDisplayValue();
            if (cur.indexOf(note) === -1) c.setValue(cur ? cur + "; " + note : note);
          }
          return out("exists");
        }
      }
    }
    sh.insertRowAfter(1);
    sh.getRange(2, 1, 1, 6).setValues([[
      isNaN(when) ? new Date() : when, name, pretty,
      '=HYPERLINK("sms:+1' + phone + '","Text")', false, note
    ]]).setFontWeight("normal").setBackground(null).setFontColor(null);
    sh.getRange(2, 1).setNumberFormat("ddd m/d h:mm am/pm");
    sh.getRange(2, 5).insertCheckboxes();
  } finally {
    lock.releaseLock();
  }
  return out("ok");
}

// One-time cleanup: merge rows that share a phone number. Keeps the oldest row,
// checks "Texted" if any copy was checked, and combines their notes.
function removeDuplicates() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const last = sh.getLastRow();
  if (last < 3) return;
  const rows = sh.getRange(2, 1, last - 1, 6).getDisplayValues();
  const keep = {}, drop = [];
  for (let i = rows.length - 1; i >= 0; i--) { // bottom = oldest
    const phone = rows[i][2];
    if (!phone) continue;
    if (!(phone in keep)) { keep[phone] = i; continue; }
    const k = keep[phone];
    if (rows[i][4] === "TRUE") sh.getRange(k + 2, 5).setValue(true);
    const note = rows[i][5], cur = sh.getRange(k + 2, 6).getDisplayValue();
    if (note && cur.indexOf(note) === -1) sh.getRange(k + 2, 6).setValue(cur ? cur + "; " + note : note);
    drop.push(i + 2);
  }
  drop.sort((a, b) => b - a).forEach(r => sh.deleteRow(r));
}

// Stop names like "=IMPORTXML(...)" from running as formulas
function clean(v) {
  const s = String(v || "").trim();
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function out(msg) {
  return ContentService.createTextOutput(JSON.stringify({ result: msg })).setMimeType(ContentService.MimeType.JSON);
}
