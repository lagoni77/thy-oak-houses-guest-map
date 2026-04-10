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

// Airbnb Gmail search query targeting both English and Danish subjects
// Using {} in Gmail query works as an OR statement.
const GMAIL_QUERY = 'from:automated@airbnb.com {subject:"Reservation confirmed" subject:"Bekræftet reservation" subject:"Ny bekræftet booking"}';

// ── MAIN SCRAPER ─────────────────────────────────────────────────────────────
/**
 * Syncs Airbnb confirmation emails into the GuestData sheet.
 * Set this as a Time-driven trigger (every hour).
 */
function syncAirbnbEmails() {
  const sheet   = getSheet();
  const existingEmailIds = getExistingEmailIds(sheet);

  // Use PropertiesService to remember the last scan time for efficiency.
  const props       = PropertiesService.getScriptProperties();
  const lastScanStr = props.getProperty('LAST_SCAN_DATE');
  const queryWithDate = lastScanStr
    ? `${GMAIL_QUERY} after:${lastScanStr}`
    : `${GMAIL_QUERY} newer_than:2m`; // 2 months back on first run

  Logger.log(`Searching Gmail with: ${queryWithDate}`);

  let threads;
  try {
    threads = GmailApp.search(queryWithDate);
  } catch (e) {
    Logger.log(`Error searching Gmail: ${e}`);
    return;
  }

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

      const body    = message.getPlainBody();
      const subject = message.getSubject();
      const parsed  = parseAirbnbEmail(body, subject, new Date(message.getDate()));

      if (!parsed) {
        Logger.log(`Could not parse email ${emailId} (Subject: ${subject}) — skipping`);
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
 * Extracts guest name, check-in/out, and home location from an Airbnb email.
 */
function parseAirbnbEmail(body, subject, emailDate) {
  try {
    let name = null;
    let location = null;
    let checkInDate = null;
    let checkOutDate = null;

    // ── GUEST NAME ──────────────────────────────────────────────────────────
    // 1. Try Danish subject format: "Bekræftet reservation – Mathias Kaisner ankommer..."
    let match = subject.match(/Bekræftet reservation\s*–\s*(.*?)\s+ankommer/i);
    if (match) name = match[1].trim();

    if (!name) {
      // 2. Try English generic ones
      const namePatterns = [
        /(?:reservation from|booked by|guest[:\s]+)\s*([A-ZæøåÆØÅ][a-zæøåÆØÅ]+(?: [A-ZæøåÆØÅ][a-zæøåÆØÅ]+)*)/i,
        /^([A-ZæøåÆØÅ][a-zæøåÆØÅ]+(?: [A-ZæøåÆØÅ][a-zæøåÆØÅ]+)*) is coming/im,
        /Hello,\s*([A-ZæøåÆØÅ][a-zæøåÆØÅ]+(?: [A-ZæøåÆØÅ][a-zæøåÆØÅ]+)*)/i
      ];
      for (const pattern of namePatterns) {
        match = body.match(pattern);
        if (match) { name = match[1].trim(); break; }
      }
    }

    if (!name) {
      // 3. Try finding name above "Identitet bekræftet" in Danish emails
      match = body.match(/([^\n]+)\s*\n+\s*Identitet bekræftet/i);
      if (match) name = match[1].trim();
    }
    
    if (!name) name = 'Unknown Guest';

    // ── LOCATION ────────────────────────────────────────────────────────────
    // 1. In Danish emails, location is directly below "Identitet bekræftet..."
    match = body.match(/Identitet bekræftet[^\n]*\s*\n+\s*([^\n]+)/i);
    if (match) location = match[1].trim();

    if (!location) {
      // 2. Try English patterns
      const locPatterns = [
        /(?:From|Home|Location|Lives in|from)\s*:\s*([A-Za-zæøåÆØÅ\s,-]+?)(?:\n|\.)/i,
        /(?:From|Home):\s*\n\s*([A-Za-zæøåÆØÅ\s,-]+?)(?:\n|$)/im
      ];
      for (const pattern of locPatterns) {
        match = body.match(pattern);
        if (match) { location = match[1].trim(); break; }
      }
    }
    
    // Fallback cleanup if the scraped location is something random like an Airbnb text
    if (!location || location.length > 50 || location.includes('Hej Brian')) {
      location = 'Unknown';
    }

    // ── DATES ───────────────────────────────────────────────────────────────
    // Parse Danish Dates specifically
    // Pattern looks for: "Indtjekning \n\n tirs. 7. jul."
    const inMatch  = body.match(/Indtjekning[^\n]*\s*\n+\s*(?:[a-zæøå]+\.?\s*)?(\d{1,2}\.\s*[a-zæøå]+\.?(?:\s*\d{4})?)/i);
    const outMatch = body.match(/Udtjekning[^\n]*\s*\n+\s*(?:[a-zæøå]+\.?\s*)?(\d{1,2}\.\s*[a-zæøå]+\.?(?:\s*\d{4})?)/i);

    if (inMatch && outMatch) {
      checkInDate  = parseDanishDate(inMatch[1], emailDate);
      checkOutDate = parseDanishDate(outMatch[1], emailDate);
    } else {
      // Fallback to English date parsing
      const datePattern = /(\b(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\.?\s+\d{1,2},?\s+\d{4}|\b\d{1,2}\s+(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{4}|\b\d{4}-\d{2}-\d{2})/gi;
      const dates = [...body.matchAll(datePattern)].map(m => new Date(m[0])).filter(d => !isNaN(d));
      if (dates.length >= 2) {
        dates.sort((a, b) => a - b);
        checkInDate  = dates[0];
        checkOutDate = dates[dates.length - 1];
      }
    }

    if (!checkInDate || !checkOutDate) {
      Logger.log('Could not find both check-in and check-out dates.');
      return null;
    }

    const checkInStr  = Utilities.formatDate(checkInDate, 'UTC', 'yyyy-MM-dd');
    const checkOutStr = Utilities.formatDate(checkOutDate, 'UTC', 'yyyy-MM-dd');

    return { name, checkIn: checkInStr, checkOut: checkOutStr, location };

  } catch (e) {
    Logger.log(`parseAirbnbEmail error: ${e.message}`);
    return null;
  }
}

/**
 * Parses Danish date strings like "7. jul." into a JS Date object.
 * Assumes the year based on the email receive date, rolling forward if the month is before email month.
 */
function parseDanishDate(dateStr, emailDate) {
  // Extract number and month
  const match = dateStr.match(/(\d{1,2})\.\s*([a-zæøå]{3,})/i);
  if (!match) return null;

  const day = parseInt(match[1]);
  const monthStr = match[2].toLowerCase();
  
  const dkMonths = {
    'jan': 0, 'feb': 1, 'mar': 2, 'apr': 3, 'maj': 4, 'jun': 5,
    'jul': 6, 'aug': 7, 'sep': 8, 'okt': 9, 'nov': 10, 'dec': 11
  };
  
  // Find which month it corresponds to
  let monthIndex = null;
  for (const [key, index] of Object.entries(dkMonths)) {
    if (monthStr.startsWith(key)) {
      monthIndex = index;
      break;
    }
  }
  
  if (monthIndex === null) return null;

  // Determine year: if booking date month is before email date month, it's probably next year.
  let year = emailDate.getFullYear();
  if (monthIndex < emailDate.getMonth() - 2) { 
    // Usually bookings aren't made 10 months past, so it must be next year
    year += 1;
  }
  
  // See if year was explicitly specified in the string (e.g. "7. jul. 2026")
  const yearMatch = dateStr.match(/\d{4}/);
  if (yearMatch) {
    year = parseInt(yearMatch[0]);
  }

  return new Date(Date.UTC(year, monthIndex, day));
}

// ── GEOCODER ─────────────────────────────────────────────────────────────────
/**
 * Returns { lat, lng } for a location string.
 * First checks existing rows in the sheet to avoid re-geocoding the same city.
 */
function getOrGeocodeLocation(location, sheet) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][COL.LOCATION - 1]).toLowerCase() === location.toLowerCase()) {
      const lat = data[i][COL.LAT - 1];
      const lng = data[i][COL.LNG - 1];
      if (lat && lng) {
        Logger.log(`Using cached coords for "${location}": ${lat}, ${lng}`);
        return { lat: parseFloat(lat), lng: parseFloat(lng) };
      }
    }
  }

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
 */
function doGet(e) {
  const sheet  = getSheet();
  const values = sheet.getDataRange().getValues();

  const guests = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (!row[COL.NAME - 1]) continue;

    guests.push({
      ID:       row[COL.ID       - 1],
      Name:     row[COL.NAME     - 1],
      CheckIn:  row[COL.CHECK_IN - 1],
      CheckOut: row[COL.CHECK_OUT- 1],
      Location: row[COL.LOCATION - 1],
      Lat:      row[COL.LAT      - 1],
      Lng:      row[COL.LNG      - 1]
    });
  }

  return ContentService
    .createTextOutput(JSON.stringify(guests))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── HELPERS ──────────────────────────────────────────────────────────────────
function getSheet() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    // If it doesn't exist, try to fall back to 'Ark1' or 'Sheet1' automatically. 
    // We'll rename it later just to be safe.
    sheet = ss.getSheets()[0];
    if (sheet && sheet.getName() !== SHEET_NAME) {
        sheet.setName(SHEET_NAME);
    }
  }
  return sheet;
}

function getExistingEmailIds(sheet) {
  const values = sheet.getDataRange().getValues();
  const ids    = new Set();
  for (let i = 1; i < values.length; i++) {
    const emailId = values[i][COL.EMAIL_ID - 1];
    if (emailId) ids.add(String(emailId));
  }
  return ids;
}

function setupSheet() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let sheet   = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.getSheets()[0];
    sheet.setName(SHEET_NAME);
    Logger.log('Renamed default sheet to: ' + SHEET_NAME);
  }

  sheet.getRange(1, 1, 1, 8).setValues([[
    'ID', 'Name', 'CheckIn', 'CheckOut', 'Location', 'Lat', 'Lng', 'EmailID'
  ]]);

  sheet.getRange(1, 1, 1, 8)
    .setBackground('#1a3327')
    .setFontColor('#4ade80')
    .setFontWeight('bold');

  sheet.setFrozenRows(1);
  Logger.log('Sheet setup complete!');
}

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
