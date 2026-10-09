// Writes _site/count.json with the number of free-haircut launch spots claimed:
// unique clients with an active Whole New You Color Visit booking created during
// the launch window. Run by the GitHub Actions deploy workflow (SQUARE_TOKEN secret).
import { writeFileSync } from "node:fs";

const LOCATION_ID = "LMF0NR8GP2YVJ";
const VARIATION_ID = "5ZFBRXTVTHPICZLJFZVLULDE";
const CREATED_FROM = new Date("2026-10-05T00:00:00-04:00");
const CREATED_TO = new Date("2026-10-31T23:59:59-04:00");
// Appointment dates to scan; Square allows at most 31 days per request.
const SCAN_FROM = new Date("2026-10-01T00:00:00-04:00");
const SCAN_WINDOWS = 4;
const ACTIVE = new Set(["ACCEPTED", "PENDING"]);
const DAY = 86400000;

async function listWindow(token, min, max) {
  const out = [];
  let cursor;
  do {
    const q = new URLSearchParams({ location_id: LOCATION_ID, start_at_min: min.toISOString(), start_at_max: max.toISOString(), limit: "100" });
    if (cursor) q.set("cursor", cursor);
    const r = await fetch(`https://connect.squareup.com/v2/bookings?${q}`, {
      headers: { Authorization: `Bearer ${token}`, "Square-Version": "2024-01-18" },
    });
    const d = await r.json();
    if (!r.ok) throw new Error(`Square ${r.status}: ${JSON.stringify(d.errors || d)}`);
    out.push(...(d.bookings || []));
    cursor = d.cursor;
  } while (cursor);
  return out;
}

async function countClaimed(token) {
  const clients = new Set();
  for (let i = 0; i < SCAN_WINDOWS; i++) {
    const min = new Date(SCAN_FROM.getTime() + i * 31 * DAY);
    const max = new Date(min.getTime() + 31 * DAY);
    for (const b of await listWindow(token, min, max)) {
      const created = new Date(b.created_at);
      const isVisit = (b.appointment_segments || []).some(s => s.service_variation_id === VARIATION_ID);
      if (isVisit && ACTIVE.has(b.status) && created >= CREATED_FROM && created <= CREATED_TO) {
        clients.add(b.customer_id || b.id);
      }
    }
  }
  return clients.size;
}

try {
  const claimed = await countClaimed(process.env.SQUARE_TOKEN);
  writeFileSync("_site/count.json", JSON.stringify({ claimed, updated: new Date().toISOString() }));
  console.log(`claimed: ${claimed}`);
} catch (e) {
  // Leave count.json out; the page falls back to its built-in number.
  console.error(e.message);
}
