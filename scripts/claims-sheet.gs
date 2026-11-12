/**
 * Google Apps Script for the "Nora - $100 claims" sheet.
 * The site POSTs each claim here (see the claim form handler in index.html);
 * this adds it as a new row at the top with an unchecked "Texted" box.
 * A phone number already in the sheet is not added twice; H1 shows the count of distinct people.
 * Every 5 minutes syncBookings() fills "Booked online" from Square for people who booked themselves,
 * and adds a row for any new client (first seen in Square since the Oct 1 launch) who booked online
 * without filling out the form
 * (needs Script Property SQUARE_TOKEN; run setupBookingCheck() once to start it).
 *
 * Setup: Extensions > Apps Script, paste this file, run setup() once,
 * then Deploy > New deployment > Web app (Execute as: Me, Who has access: Anyone).
 */

const SHEET_NAME = "Claims";
const HEADERS = ["Claimed", "Name", "Phone", "Text her", "Texted", "Notes", "Booked online"];

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME, 0);
  sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS])
    .setFontWeight("bold").setBackground("#4A2F22").setFontColor("#F3ECE3");
  sh.setFrozenRows(1);
  sh.setColumnWidth(1, 150); sh.setColumnWidth(2, 120); sh.setColumnWidth(3, 130);
  sh.setColumnWidth(4, 90); sh.setColumnWidth(5, 70); sh.setColumnWidth(6, 260); sh.setColumnWidth(7, 190);
  // Gray out rows once "Texted" is checked
  const rule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied("=$E2=TRUE").setFontColor("#9A8B80")
    .setRanges([sh.getRange("A2:G1000")]).build();
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
    // Same phone already listed: add the note to that row instead of a duplicate
    const last = sh.getLastRow();
    if (last > 1) {
      const phones = sh.getRange(2, 3, last - 1, 1).getDisplayValues();
      for (let i = 0; i < phones.length; i++) {
        if (phones[i][0] === pretty(phone)) {
          if (note) {
            const c = sh.getRange(i + 2, 6), cur = c.getDisplayValue();
            if (cur.indexOf(note) === -1) c.setValue(cur ? cur + "; " + note : note);
          }
          return out("exists");
        }
      }
    }
    addRow(sh, isNaN(when) ? new Date() : when, name, phone, note);
  } finally {
    lock.releaseLock();
  }
  saveMetaIds(phone, d);
  try { sendLead(phone, name, d); } catch (err) { console.error("CAPI Lead: " + err.message); }
  return out("ok");
}

// Adds a person as row 2 (newest at the top) with an unchecked "Texted" box
function addRow(sh, when, name, phone, note) {
  sh.insertRowAfter(1);
  sh.getRange(2, 1, 1, 6).setValues([[
    when, name, pretty(phone), '=HYPERLINK("sms:+1' + phone + '","Text")', false, note
  ]]);
  // The new row copies the header's style; reset it across every column, including "Booked online"
  sh.getRange(2, 1, 1, HEADERS.length).setFontWeight("normal").setBackground(null).setFontColor(null);
  sh.getRange(2, 1).setNumberFormat("ddd m/d h:mm am/pm");
  sh.getRange(2, 5).insertCheckboxes();
}

function pretty(phone) {
  return "(" + phone.slice(0, 3) + ") " + phone.slice(3, 6) + "-" + phone.slice(6);
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

// Run once: adds the "Booked online" column, starts the 5-minute Square check, and runs it now.
function setupBookingCheck() {
  setup();
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === "syncBookings")
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger("syncBookings").timeBased().everyMinutes(5).create();
  syncBookings();
}

// Marks each person who booked herself in Square (online booking site, Google, Instagram)
// with her upcoming or past appointment times. New clients who booked online without the form are added;
// appointments Nora enters and existing clients who book themselves are not.
function syncBookings() {
  const token = PropertiesService.getScriptProperties().getProperty("SQUARE_TOKEN");
  if (!token) throw new Error("Add the Square token as Script Property SQUARE_TOKEN (Project Settings > Script Properties).");
  const headers = { Authorization: "Bearer " + token.trim(), "Square-Version": "2025-01-23", "Content-Type": "application/json" };
  const api = (path, opts) => {
    const r = UrlFetchApp.fetch("https://connect.squareup.com/v2/" + path, Object.assign({ headers: headers, muteHttpExceptions: true }, opts || {}));
    if (r.getResponseCode() !== 200) throw new Error("Square " + r.getResponseCode() + ": " + r.getContentText().slice(0, 200));
    return JSON.parse(r.getContentText());
  };

  // Bookings come in windows of at most 31 days: from the launch to 4 months ahead
  const bookings = [];
  const end = Date.now() + 120 * 864e5;
  for (let lo = Date.parse("2026-10-01T00:00:00Z"); lo < end; lo += 30 * 864e5) {
    let cursor = "";
    do {
      const q = "bookings?location_id=LMF0NR8GP2YVJ&limit=100&start_at_min=" + new Date(lo).toISOString() +
        "&start_at_max=" + new Date(lo + 30 * 864e5).toISOString() + (cursor ? "&cursor=" + encodeURIComponent(cursor) : "");
      const d = api(q);
      (d.bookings || []).forEach(b => bookings.push(b));
      cursor = d.cursor || "";
    } while (cursor);
  }
  const self = bookings.filter(b =>
    (b.source === "FIRST_PARTY_BUYER" || b.source === "THIRD_PARTY_BUYER") &&
    (b.status === "ACCEPTED" || b.status === "PENDING") && b.customer_id);

  // Service names, to tell Nora what someone without a claim form booked
  const serviceOf = {};
  let cur = "";
  do {
    const d = api("catalog/list?types=ITEM" + (cur ? "&cursor=" + encodeURIComponent(cur) : ""));
    (d.objects || []).forEach(o => (o.item_data.variations || []).forEach(v => serviceOf[v.id] = o.item_data.name));
    cur = d.cursor || "";
  } while (cur);

  // Customer names and phone numbers, 100 at a time
  const LAUNCH = "2026-10-01";
  const phoneOf = {}, nameOf = {}, isNew = {};
  const ids = [...new Set(self.map(b => b.customer_id))];
  for (let i = 0; i < ids.length; i += 100) {
    const d = api("customers/bulk-retrieve", { method: "post", payload: JSON.stringify({ customer_ids: ids.slice(i, i + 100) }) });
    Object.keys(d.responses || {}).forEach(id => {
      const c = (d.responses[id] || {}).customer;
      if (c && c.phone_number) phoneOf[id] = c.phone_number.replace(/\D/g, "").slice(-10);
      if (c) isNew[id] = (c.created_at || "") >= LAUNCH;
      if (c) nameOf[id] = clean(((c.given_name || "") + " " + (c.family_name || "").slice(0, 1)).trim()).slice(0, 40);
    });
  }
  const byPhone = {}, firstBooking = {};
  self.sort((a, b) => a.start_at < b.start_at ? -1 : 1).forEach(b => {
    const ph = phoneOf[b.customer_id];
    if (!ph || ph.length !== 10) return;
    const t = Utilities.formatDate(new Date(b.start_at), "America/New_York", "EEE M/d h:mm a");
    (byPhone[ph] = byPhone[ph] || []).push(t);
    if (isNew[b.customer_id] && (!firstBooking[ph] || b.created_at < firstBooking[ph].created_at)) firstBooking[ph] = b;
  });

  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (sh.getRange(1, 7).getValue() !== "Booked online") setup();
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    // Add new clients who booked online but never filled out the form, oldest first so the newest ends up on top
    const listed = new Set(sh.getLastRow() > 1 ? sh.getRange(2, 3, sh.getLastRow() - 1, 1).getDisplayValues().map(r => r[0].replace(/\D/g, "")) : []);
    Object.keys(firstBooking).filter(ph => !listed.has(ph))
      .sort((a, b) => firstBooking[a].created_at < firstBooking[b].created_at ? -1 : 1)
      .forEach(ph => {
        const b = firstBooking[ph];
        const what = (b.appointment_segments || []).map(s => serviceOf[s.service_variation_id] || "a service").join(" + ");
        addRow(sh, new Date(b.created_at), nameOf[b.customer_id] || "(no name)", ph, "No claim form; booked " + what + " online");
      });
    if (sh.getLastRow() < 2) return;
    const n = sh.getLastRow() - 1; // re-read under the lock in case a claim just came in
    const phones = sh.getRange(2, 3, n, 1).getDisplayValues();
    const vals = phones.map(r => {
      const appts = byPhone[r[0].replace(/\D/g, "")];
      return [appts ? "Yes: " + appts.join(", ") : ""];
    });
    sh.getRange(2, 7, n, 1).setValues(vals);
  } finally {
    lock.releaseLock();
  }
}

/* ===================== Meta Conversions API =====================
 * Tells Meta which leads turned into paying clients, so ads optimize for buyers, not form fills.
 *   Lead      sent the moment someone claims (same event_id as the browser pixel, so Meta counts it once)
 *   Schedule  a Square booking for someone in the sheet
 *   Purchase  a completed Square payment for someone in the sheet, with the real amount (refunds subtracted)
 * Matching uses hashed phone, email and first name from Square, plus the Meta click/browser IDs
 * the claim form saves (Script Properties, so no sheet columns are touched).
 * Script Properties: META_TOKEN (required), META_TEST_CODE (optional: events show only under
 * Events Manager > Test events and don't affect ads; delete it to go live).
 * Run setupConversions() once: sends everything since the Oct 1 launch, then checks hourly.
 * Each event is sent once; sent IDs are kept in Script Property CAPI_SENT.
 */
const PIXEL_ID = "2106667673542937";
const CAPI_FROM = "2026-10-01T00:00:00Z";

function setupConversions() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === "sendConversions")
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger("sendConversions").timeBased().everyHours(1).create();
  sendConversions();
}

function sendConversions() {
  const props = PropertiesService.getScriptProperties();
  const sq = squareApi();
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (sh.getLastRow() < 2) return;
  const inSheet = new Set(sh.getRange(2, 3, sh.getLastRow() - 1, 1).getDisplayValues().map(r => r[0].replace(/\D/g, "")).filter(p => p.length === 10));

  const payments = [];
  let cur = "";
  do {
    const d = sq("payments?location_id=LMF0NR8GP2YVJ&limit=100&begin_time=" + CAPI_FROM + (cur ? "&cursor=" + encodeURIComponent(cur) : ""));
    (d.payments || []).filter(p => p.status === "COMPLETED" && p.customer_id).forEach(p => payments.push(p));
    cur = d.cursor || "";
  } while (cur);

  const bookings = [];
  for (let lo = Date.parse(CAPI_FROM), end = Date.now() + 120 * 864e5; lo < end; lo += 30 * 864e5) {
    let c = "";
    do {
      const d = sq("bookings?location_id=LMF0NR8GP2YVJ&limit=100&start_at_min=" + new Date(lo).toISOString() +
        "&start_at_max=" + new Date(lo + 30 * 864e5).toISOString() + (c ? "&cursor=" + encodeURIComponent(c) : ""));
      (d.bookings || []).filter(b => b.customer_id && (b.status === "ACCEPTED" || b.status === "PENDING")).forEach(b => bookings.push(b));
      c = d.cursor || "";
    } while (c);
  }

  const customers = {};
  const ids = [...new Set(payments.map(p => p.customer_id).concat(bookings.map(b => b.customer_id)))];
  for (let i = 0; i < ids.length; i += 100) {
    const d = sq("customers/bulk-retrieve", { method: "post", payload: JSON.stringify({ customer_ids: ids.slice(i, i + 100) }) });
    Object.keys(d.responses || {}).forEach(id => { const c = (d.responses[id] || {}).customer; if (c) customers[id] = c; });
  }
  const phoneOf = id => ((customers[id] || {}).phone_number || "").replace(/\D/g, "").slice(-10);

  const sent = new Set(JSON.parse(props.getProperty("CAPI_SENT") || "[]"));
  const events = [];
  payments.forEach(p => {
    const ph = phoneOf(p.customer_id), id = "sq_pay_" + p.id;
    if (!inSheet.has(ph) || sent.has(id)) return;
    const cents = p.amount_money.amount - ((p.refunded_money || {}).amount || 0);
    if (cents <= 0) return;
    events.push(capiEvent("Purchase", id, p.created_at, ph, customers[p.customer_id], { currency: "USD", value: cents / 100 }));
  });
  bookings.forEach(b => {
    const ph = phoneOf(b.customer_id), id = "sq_book_" + b.id;
    if (!inSheet.has(ph) || sent.has(id)) return;
    events.push(capiEvent("Schedule", id, b.created_at, ph, customers[b.customer_id], {}));
  });
  // Meta rejects events older than 62 days (offline); drop those instead of failing the batch
  const oldest = Date.now() - 61 * 864e5;
  const fresh = events.filter(e => e.event_time * 1000 > oldest);
  for (let i = 0; i < fresh.length; i += 500) {
    const batch = fresh.slice(i, i + 500);
    postCapi(batch);
    if (!props.getProperty("META_TEST_CODE")) batch.forEach(e => sent.add(e.event_id)); // test events don't count, so they must be sent again live
  }
  props.setProperty("CAPI_SENT", JSON.stringify([...sent].slice(-300)));
  console.log("CAPI sent " + fresh.length + " events");
}

function capiEvent(name, eventId, isoTime, phone, customer, custom) {
  const c = customer || {};
  const meta = JSON.parse(PropertiesService.getScriptProperties().getProperty("meta:" + phone) || "{}");
  const ud = { ph: [sha("1" + phone)], country: [sha("us")] };
  if (c.email_address) ud.em = [sha(c.email_address.trim().toLowerCase())];
  if (c.given_name) ud.fn = [sha(c.given_name.trim().toLowerCase())];
  if (c.family_name) ud.ln = [sha(c.family_name.trim().toLowerCase())];
  if (meta.fbc) ud.fbc = meta.fbc;
  if (meta.fbp) ud.fbp = meta.fbp;
  return { event_name: name, event_time: Math.floor(Date.parse(isoTime) / 1000), event_id: eventId,
    action_source: "physical_store", user_data: ud, custom_data: custom };
}

// Server copy of the browser Lead; same event_id, so Meta keeps one
function sendLead(phone, name, d) {
  if (!d.event_id) return;
  const ud = { ph: [sha("1" + phone)], fn: [sha(String(name).replace(/^'/, "").trim().toLowerCase())], country: [sha("us")] };
  if (d.fbc) ud.fbc = String(d.fbc).slice(0, 300);
  if (d.fbp) ud.fbp = String(d.fbp).slice(0, 100);
  if (d.ua) ud.client_user_agent = String(d.ua).slice(0, 400);
  postCapi([{ event_name: "Lead", event_time: Math.floor(Date.now() / 1000), event_id: String(d.event_id).slice(0, 60),
    action_source: "website", event_source_url: "https://norascotto.com/", user_data: ud }]);
}

// Keeps the Meta click (fbc) and browser (fbp) IDs from the claim form, matched to the phone, for later Purchase events
function saveMetaIds(phone, d) {
  if (!d.fbc && !d.fbp) return;
  const props = PropertiesService.getScriptProperties();
  const old = JSON.parse(props.getProperty("meta:" + phone) || "{}");
  props.setProperty("meta:" + phone, JSON.stringify({
    fbc: d.fbc ? String(d.fbc).slice(0, 300) : old.fbc, fbp: d.fbp ? String(d.fbp).slice(0, 100) : old.fbp }));
}

function postCapi(events) {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty("META_TOKEN");
  if (!token) throw new Error("Add the Meta token as Script Property META_TOKEN.");
  const body = { data: events, access_token: token.trim() };
  const test = props.getProperty("META_TEST_CODE");
  if (test) body.test_event_code = test.trim();
  const r = UrlFetchApp.fetch("https://graph.facebook.com/v21.0/" + PIXEL_ID + "/events",
    { method: "post", contentType: "application/json", payload: JSON.stringify(body), muteHttpExceptions: true });
  if (r.getResponseCode() !== 200) throw new Error("Meta " + r.getResponseCode() + ": " + r.getContentText().slice(0, 300));
  return JSON.parse(r.getContentText());
}

function squareApi() {
  const token = PropertiesService.getScriptProperties().getProperty("SQUARE_TOKEN");
  if (!token) throw new Error("Add the Square token as Script Property SQUARE_TOKEN.");
  const headers = { Authorization: "Bearer " + token.trim(), "Square-Version": "2025-01-23", "Content-Type": "application/json" };
  return (path, opts) => {
    const r = UrlFetchApp.fetch("https://connect.squareup.com/v2/" + path, Object.assign({ headers: headers, muteHttpExceptions: true }, opts || {}));
    if (r.getResponseCode() !== 200) throw new Error("Square " + r.getResponseCode() + ": " + r.getContentText().slice(0, 200));
    return JSON.parse(r.getContentText());
  };
}

function sha(s) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8)
    .map(b => ("0" + (b & 255).toString(16)).slice(-2)).join("");
}

// Stop names like "=IMPORTXML(...)" from running as formulas
function clean(v) {
  const s = String(v || "").trim();
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function out(msg) {
  return ContentService.createTextOutput(JSON.stringify({ result: msg })).setMimeType(ContentService.MimeType.JSON);
}
