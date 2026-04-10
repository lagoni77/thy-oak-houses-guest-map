// ═══════════════════════════════════════════════════════════════════════════════
// THY OAK HOUSES — AIRBNB GUEST MAP BACKEND
// Paste this entire file into: Google Sheets → Extensions → Apps Script
//
// SETUP CHECKLIST:
//   1. Create a sheet named "GuestData" with columns:
//      ID | Name | CheckIn | CheckOut | Location | Lat | Lng | EmailID
//   2. Enable the Gmail API:
//      Left sidebar → Services (+) → Gmail API → Add
//   3. Deploy as Web App:
//      Deploy → New Deployment → Web App
//      Execute as: Me | Who has access: Anyone
//   4. Set a Time-driven Trigger (Clock icon on left):
//      Function: syncAirbnbEmails | Time-based: Every hour
// ═══════════════════════════════════════════════════════════════════════════════

// ── CONFIG ──────────────────────────────────────────────────────────────────
const SHEET_NAME = 'GuestData';

// Column indices (1-based) — must match your sheet header order
const COL = {
  ID:       1,
  NAME:     2,
  CHECK_IN: 3,
  CHECK_OUT:4,
  LOCATION: 5,
  LAT:      6,
  LNG:      7,
  EMAIL_ID: 8
};

// Airbnb Gmail search query
const GMAIL_QUERY = 'from:automated@airbnb.com subject:"Reservation confirmed"';

// ── MAIN SCRAPER ─────────────────────────────────────────────────────────────
/**
 * Syncs Airbnb confirmation emails into the GuestData sheet.
 * Set this as a Time-driven trigger (every hour).
 */
function syncAirbnbEmails() {
  const sheet   = getSheet();
  const existingEmailIds = getExistingEmailIds(sheet);

  // Use PropertiesService to remember the last scan time for efficiency.
  // On first run this will be null, so we scan all mail.
  const props       = PropertiesService.getScriptProperties();
  const lastScanStr = props.getProperty('LAST_SCAN_DATE');
  const queryWithDate = lastScanStr
    ? `${GMAIL_QUERY} after:${lastScanStr}`
    : GMAIL_QUERY;

  Logger.log(`Searching Gmail with: ${queryWithDate}`);

  const threads = GmailApp.search(queryWithDate);
  Logger.log(`Found ${threads.length} thread(s)`);

  let newRowsCount = 0;

  threads.forEach(thread => {
    thread.getMessages().forEach(message => {
      const emailId = message.getId();

      // Skip if we've already processed this email
      if (existingEmailIds.has(emailId)) {
        Logger.log(`Skipping already-processed email: ${emailId}`);
        return;
      }

      const body = message.getPlainBody();
      const parsed = parseAirbnbEmail(body);

      if (!parsed) {
        Logger.log(`Could not parse email ${emailId} — skipping`);
        return;
      }

      // Geocode only if location is new (saves API quota)
      const coords = getOrGeocodeLocation(parsed.location, sheet);

      // Generate unique row ID
      const rowId = `TOH-${emailId.substring(0, 8).toUpperCase()}`;

      sheet.appendRow([
        rowId,
        parsed.name,
        parsed.checkIn,
        parsed.checkOut,
        parsed.location,
        coords.lat,
        coords.lng,
        emailId
      ]);

      existingEmailIds.add(emailId);
      newRowsCount++;
      Logger.log(`Added guest: ${parsed.name} from ${parsed.location}`);
    });
  });

  // Save the current date so next run is incremental
  const today = Utilities.formatDate(new Date(), 'UTC', 'yyyy/MM/dd');
  props.setProperty('LAST_SCAN_DATE', today);

  Logger.log(`Sync complete. Added ${newRowsCount} new guest(s).`);
}

// ── EMAIL PARSER ─────────────────────────────────────────────────────────────
/**
 * Extracts guest name, check-in/out, and home location from an Airbnb
 * confirmation email body.
 * Returns { name, checkIn, checkOut, location } or null if parsing fails.
 */
function parseAirbnbEmail(body) {
  try {
    // ── Guest Name ──────────────────────────────────────────────────────────
    // Airbnb emails typically say "You have a new reservation from [Name]"
    // or "[Name] is coming to stay"
    const namePatterns = [
      /(?:reservation from|booked by|guest[:\s]+)\s*([A-Z][a-z]+(?: [A-Z][a-z]+)*)/i,
      /^([A-Z][a-z]+(?: [A-Z][a-z]+)*) is coming/im,
      /Hello,\s*([A-Z][a-z]+(?: [A-Z][a-z]+)*)/i,
      /Guest:\s*([A-Za-z ]+)/i
    ];

    let name = null;
    for (const pattern of namePatterns) {
      const match = body.match(pattern);
      if (match) { name = match[1].trim(); break; }
    }
    if (!name) name = 'Unknown Guest';

    // ── Dates ───────────────────────────────────────────────────────────────
    // Match patterns like "Jun 15, 2024" or "15 June 2024" or "2024-06-15"
    const datePattern = /(\b(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\.?\s+\d{1,2},?\s+\d{4}|\b\d{1,2}\s+(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{4}|\b\d{4}-\d{2}-\d{2})/gi;
    const dates = [...body.matchAll(datePattern)].map(m => new Date(m[0])).filter(d => !isNaN(d));

    if (dates.length < 2) {
      Logger.log('Could not find 2 dates — skipping');
      return null;
    }

    dates.sort((a, b) => a - b);
    const checkIn  = Utilities.formatDate(dates[0], 'UTC', 'yyyy-MM-dd');
    const checkOut = Utilities.formatDate(dates[dates.length - 1], 'UTC', 'yyyy-MM-dd');

    // ── Location ────────────────────────────────────────────────────────────
    // Airbnb includes "Home" location (guest's city/country)
    const locationPatterns = [
      /(?:From|Home|Location|Lives in|from)\s*:\s*([A-Za-z\s,]+?)(?:\n|\.)/i,
      /(?:From|Home):\s*\n\s*([A-Za-z\s,]+?)(?:\n|$)/im,
      /(\b[A-Z][a-z]+(?:,\s*[A-Z][a-z]+)*(?:,\s*[A-Z]{2,3})?)/  // City, Country
    ];

    let location = 'Unknown';
    for (const pattern of locationPatterns) {
      const match = body.match(pattern);
      if (match) { location = match[1].trim(); break; }
    }

    return { name, checkIn, checkOut, location };

  } catch (e) {
    Logger.log(`parseAirbnbEmail error: ${e.message}`);
    return null;
  }
}

// ── GEOCODER ─────────────────────────────────────────────────────────────────
/**
 * Returns { lat, lng } for a location string.
 * First checks existing rows in the sheet to avoid re-geocoding the same city.
 */
function getOrGeocodeLocation(location, sheet) {
  // Check if we already have coords for this location
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {  // start at 1 to skip header
    if (String(data[i][COL.LOCATION - 1]).toLowerCase() === location.toLowerCase()) {
      const lat = data[i][COL.LAT - 1];
      const lng = data[i][COL.LNG - 1];
      if (lat && lng) {
        Logger.log(`Using cached coords for "${location}": ${lat}, ${lng}`);
        return { lat: parseFloat(lat), lng: parseFloat(lng) };
      }
    }
  }

  // Not found — geocode it
  try {
    const geo     = Maps.newGeocoder().geocode(location);
    const results = geo.results;
    if (results && results.length > 0) {
      const coords = results[0].geometry.location;
      Logger.log(`Geocoded "${location}": ${coords.lat}, ${coords.lng}`);
      return { lat: coords.lat, lng: coords.lng };
    }
  } catch (e) {
    Logger.log(`Geocoding failed for "${location}": ${e.message}`);
  }

  Logger.log(`No coords found for "${location}" — using 0,0`);
  return { lat: 0, lng: 0 };
}

// ── API ENDPOINT ─────────────────────────────────────────────────────────────
/**
 * HTTP GET handler — deployed as Web App.
 * Returns all guest data as a JSON array with CORS headers.
 * URL: Deploy > New Deployment > Web App > "Who has access: Anyone"
 */
function doGet(e) {
  const sheet  = getSheet();
  const values = sheet.getDataRange().getValues();

  // Build array from rows (skip header row)
  const guests = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (!row[COL.NAME - 1]) continue;  // skip empty rows

    guests.push({
      ID:       row[COL.ID       - 1],
      Name:     row[COL.NAME     - 1],
      CheckIn:  row[COL.CHECK_IN - 1],
      CheckOut: row[COL.CHECK_OUT- 1],
      Location: row[COL.LOCATION - 1],
      Lat:      row[COL.LAT      - 1],
      Lng:      row[COL.LNG      - 1]
      // Note: EmailID intentionally excluded from public API
    });
  }

  const json = JSON.stringify(guests);

  // Return with CORS headers so thy-oak-houses.vercel.app can fetch
  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

// ── HELPERS ──────────────────────────────────────────────────────────────────
/**
 * Returns the GuestData sheet. Throws if it doesn't exist.
 */
function getSheet() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error(`Sheet "${SHEET_NAME}" not found. Please create it first.`);
  return sheet;
}

/**
 * Returns a Set of all EmailIDs already in the sheet (for duplicate detection).
 */
function getExistingEmailIds(sheet) {
  const values = sheet.getDataRange().getValues();
  const ids    = new Set();
  for (let i = 1; i < values.length; i++) {
    const emailId = values[i][COL.EMAIL_ID - 1];
    if (emailId) ids.add(String(emailId));
  }
  return ids;
}

/**
 * Utility: Run this once manually to create the sheet header row.
 */
function setupSheet() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let sheet   = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    Logger.log('Created sheet: ' + SHEET_NAME);
  }

  // Set header row
  sheet.getRange(1, 1, 1, 8).setValues([[
    'ID', 'Name', 'CheckIn', 'CheckOut', 'Location', 'Lat', 'Lng', 'EmailID'
  ]]);

  // Style the header
  sheet.getRange(1, 1, 1, 8)
    .setBackground('#1a3327')
    .setFontColor('#4ade80')
    .setFontWeight('bold');

  sheet.setFrozenRows(1);
  Logger.log('Sheet setup complete!');
}

// ── MANUAL TEST ──────────────────────────────────────────────────────────────
/**
 * Run this manually to test with a single fake guest entry.
 */
function testAddSampleGuest() {
  const sheet = getSheet();
  const coords = getOrGeocodeLocation('Copenhagen, Denmark', sheet);
  sheet.appendRow([
    'TOH-TEST001',
    'Test Guest',
    '2025-06-01',
    '2025-06-07',
    'Copenhagen, Denmark',
    coords.lat,
    coords.lng,
    'TEST-EMAIL-ID-001'
  ]);
  Logger.log('Sample guest added!');
}
