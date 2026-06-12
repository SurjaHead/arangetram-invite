/*
  Paste this file into the Google Apps Script project behind the RSVP web app.

  If the script is attached to the RSVP spreadsheet, leave SPREADSHEET_ID blank.
  If it is a standalone Apps Script project, paste the spreadsheet ID below.

  After saving:
  1. Run installReminderTrigger once and approve permissions.
  2. Deploy > Manage deployments > Edit web app > New version > Deploy.
*/

const SPREADSHEET_ID = "";

const CONFIG = {
  sheetName: "Responses",
  eventTitle: "Aadhya & Thanvi's Bharatanatyam Ranga Pravesham",
  eventDateLabel: "Saturday, July 25, 2026",
  eventTimeLabel: "From 3:00pm Onwards",
  doorsLabel: "Doors open at 2:30 PM",
  eventStartUtc: "20260725T190000Z",
  eventEndUtc: "20260725T220000Z",
  eventLocation: "Shenkman Arts Centre, 245 Centrum Blvd., Ottawa, ON K1E 0A1",
  siteUrl: "https://aadhyathanviarangetram.com/",
  googleCalendarUrl: "https://calendar.google.com/calendar/render?action=TEMPLATE&text=Aadhya%20%26%20Thanvi%27s%20Bharatanatyam%20Ranga%20Pravesham&dates=20260725T150000%2F20260725T180000&ctz=America%2FToronto&details=Doors%20open%20at%202%3A30%20PM.%20You%27re%20invited%20to%20Aadhya%20Sure%20and%20Thanvi%20Vashishta%20Nagineni%27s%20Bharatanatyam%20Ranga%20Pravesham.&location=Shenkman%20Arts%20Centre%2C%20245%20Centrum%20Blvd.%2C%20Ottawa%2C%20ON%20K1E%200A1",
  icsUrl: "https://aadhyathanviarangetram.com/aadhya-thanvi-ranga-pravesham.ics",
  reminderAt: "2026-07-23T10:00:00-04:00"
};

const HEADERS = [
  "Timestamp",
  "Name",
  "Email",
  "Phone",
  "Status",
  "Adults",
  "Children",
  "Message",
  "Submitted At",
  "Confirmation Sent At",
  "Reminder Sent At"
];

const HEADER_ALIASES = {
  "Name": ["Full Name"],
  "Email": ["Email Address", "E-mail"],
  "Phone": ["Phone Number", "Mobile"],
  "Status": ["Attendance", "RSVP"],
  "Adults": ["Adult Count", "Number of Adults"],
  "Children": ["Child Count", "Kids", "Number of Children"],
  "Message": ["Message for Aadhya & Thanvi", "Notes", "Wishes"],
  "Submitted At": ["Date", "Submitted", "Submission Date"]
};

function doGet(e) {
  const params = e && e.parameter ? e.parameter : {};

  if (params.action === "comments") {
    return publicCommentsResponse_(params);
  }

  return params.callback
    ? jsonp_(params.callback, { ok: true, message: "RSVP endpoint is running." })
    : json_({ ok: true, message: "RSVP endpoint is running." });
}

function doPost(e) {
  const data = parseRequest_(e);
  data.status = normaliseStatus_(data.status);

  if (!data.name || !data.email || !data.status) {
    return json_({ ok: false, error: "Name, email, and status are required." });
  }

  const sheet = getSheet_();
  const headerMap = ensureHeaders_(sheet);
  const rowNumber = upsertRsvp_(sheet, headerMap, data);

  try {
    sendConfirmationEmail_(data);
    sheet.getRange(rowNumber, headerMap["Confirmation Sent At"]).setValue(new Date());
  } catch (error) {
    console.error("Confirmation email failed: " + error.message);
  }

  return json_({ ok: true });
}

function installReminderTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(trigger => trigger.getHandlerFunction() === "sendReminderEmails")
    .forEach(trigger => ScriptApp.deleteTrigger(trigger));

  ScriptApp.newTrigger("sendReminderEmails")
    .timeBased()
    .at(new Date(CONFIG.reminderAt))
    .create();
}

function sendReminderEmails() {
  const sheet = getSheet_();
  const headerMap = ensureHeaders_(sheet);
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) return;

  const values = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  values.forEach((row, index) => {
    const rowNumber = index + 2;
    const status = normaliseStatus_(row[headerMap.Status - 1]);
    const email = String(row[headerMap.Email - 1] || "").trim();
    const reminderSentAt = row[headerMap["Reminder Sent At"] - 1];

    if (status !== "yes" || !email || reminderSentAt) return;

    const data = {
      name: row[headerMap.Name - 1],
      email,
      status,
      adults: row[headerMap.Adults - 1],
      children: row[headerMap.Children - 1]
    };

    try {
      sendReminderEmail_(data);
      sheet.getRange(rowNumber, headerMap["Reminder Sent At"]).setValue(new Date());
    } catch (error) {
      console.error("Reminder email failed for " + email + ": " + error.message);
    }
  });
}

function dedupeExistingRsvps() {
  const sheet = getSheet_();
  const headerMap = ensureHeaders_(sheet);
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();

  if (!headerMap.Email || lastRow < 3) return "No duplicate RSVP rows found.";

  const values = sheet.getRange(2, 1, lastRow - 1, lastColumn).getValues();
  const latestByEmail = {};

  values.forEach((row, index) => {
    const email = String(row[headerMap.Email - 1] || "").trim().toLowerCase();
    if (!email) return;

    const rowNumber = index + 2;
    const previous = latestByEmail[email];
    if (!previous || rowDateValue_(row, headerMap) >= rowDateValue_(previous.row, headerMap)) {
      latestByEmail[email] = { rowNumber, row };
    }
  });

  const rowsToKeep = Object.keys(latestByEmail).reduce((memo, email) => {
    memo[latestByEmail[email].rowNumber] = true;
    return memo;
  }, {});
  const rowsToDelete = [];

  values.forEach((row, index) => {
    const rowNumber = index + 2;
    const email = String(row[headerMap.Email - 1] || "").trim().toLowerCase();
    if (email && !rowsToKeep[rowNumber]) rowsToDelete.push(rowNumber);
  });

  rowsToDelete.sort((a, b) => b - a).forEach(rowNumber => sheet.deleteRow(rowNumber));
  return "Removed " + rowsToDelete.length + " duplicate RSVP row(s).";
}

function publicCommentsResponse_(params) {
  const requestedLimit = Number(params.limit || 24);
  const limit = Number.isFinite(requestedLimit) ? Math.max(1, Math.min(requestedLimit, 60)) : 24;
  const payload = {
    ok: true,
    messages: getPublicMessages_(limit)
  };

  return params.callback ? jsonp_(params.callback, payload) : json_(payload);
}

function getPublicMessages_(limit) {
  const sheet = getSheet_();
  const headerMap = ensureHeaders_(sheet);
  const lastRow = sheet.getLastRow();

  if (!headerMap.Message || lastRow < 2) return [];

  return sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn())
    .getValues()
    .map(row => {
      const message = cleanPublicText_(row[headerMap.Message - 1], 700);
      if (!message) return null;

      const submittedAt = publicDate_(row, headerMap);
      return {
        name: publicDisplayName_(row[headerMap.Name - 1]),
        message,
        submittedAt
      };
    })
    .filter(Boolean)
    .sort((a, b) => (Date.parse(b.submittedAt) || 0) - (Date.parse(a.submittedAt) || 0))
    .slice(0, limit);
}

function parseRequest_(e) {
  const raw = e && e.postData && e.postData.contents ? e.postData.contents : "{}";

  try {
    return JSON.parse(raw);
  } catch (error) {
    return raw.split("&").reduce((memo, pair) => {
      const parts = pair.split("=");
      if (!parts[0]) return memo;
      memo[decodeURIComponent(parts[0])] = decodeURIComponent((parts[1] || "").replace(/\+/g, " "));
      return memo;
    }, {});
  }
}

function rowDateValue_(row, headerMap) {
  const submitted = row[headerMap["Submitted At"] - 1];
  const timestamp = row[headerMap.Timestamp - 1];
  const value = submitted || timestamp;
  const date = value instanceof Date ? value : new Date(value);
  const time = date.getTime();
  return isNaN(time) ? 0 : time;
}

function getSheet_() {
  const scriptPropertyId = PropertiesService.getScriptProperties().getProperty("SPREADSHEET_ID");
  const spreadsheetId = scriptPropertyId || SPREADSHEET_ID;
  const spreadsheet = spreadsheetId
    ? SpreadsheetApp.openById(spreadsheetId)
    : SpreadsheetApp.getActiveSpreadsheet();

  if (!spreadsheet) {
    throw new Error("No spreadsheet found. Add a SPREADSHEET_ID or attach this script to the RSVP spreadsheet.");
  }

  return spreadsheet.getSheetByName(CONFIG.sheetName) || spreadsheet.getSheets()[0] || spreadsheet.insertSheet(CONFIG.sheetName);
}

function ensureHeaders_(sheet) {
  const lastColumn = Math.max(sheet.getLastColumn(), HEADERS.length);
  const existing = sheet.getRange(1, 1, 1, lastColumn).getValues()[0].map(value => String(value || "").trim());
  const hasHeaders = existing.some(Boolean);

  if (!hasHeaders) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.setFrozenRows(1);
    return getHeaderMap_(sheet);
  }

  const canonicalExisting = existing.map(header => canonicalHeader_(header)).filter(Boolean);
  const missing = HEADERS.filter(header => !canonicalExisting.includes(header));
  if (missing.length) {
    const lastUsedColumn = existing.reduce((last, value, index) => value ? index + 1 : last, 0);
    sheet.getRange(1, lastUsedColumn + 1, 1, missing.length).setValues([missing]);
  }

  sheet.setFrozenRows(1);
  return getHeaderMap_(sheet);
}

function getHeaderMap_(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  return headers.reduce((map, header, index) => {
    const label = String(header || "").trim();
    const canonical = canonicalHeader_(label);
    if (canonical && !map[canonical]) map[canonical] = index + 1;
    if (label && !map[label]) map[label] = index + 1;
    return map;
  }, {});
}

function canonicalHeader_(header) {
  const label = String(header || "").trim();
  if (HEADERS.includes(label)) return label;

  const lower = label.toLowerCase();
  const exactCaseInsensitive = HEADERS.find(candidate => candidate.toLowerCase() === lower);
  if (exactCaseInsensitive) return exactCaseInsensitive;

  const canonical = HEADERS.find(candidate =>
    (HEADER_ALIASES[candidate] || []).some(alias => alias.toLowerCase() === lower)
  );

  return canonical || "";
}

function upsertRsvp_(sheet, headerMap, data) {
  const email = String(data.email || "").trim().toLowerCase();
  const matchingRows = findRowsByEmail_(sheet, headerMap, email);
  const existing = matchingRows.length
    ? sheet.getRange(matchingRows[0], 1, 1, sheet.getLastColumn()).getValues()[0]
    : new Array(sheet.getLastColumn()).fill("");
  const now = new Date();

  matchingRows.sort((a, b) => b - a).forEach(rowNumber => sheet.deleteRow(rowNumber));
  const rowNumber = sheet.getLastRow() + 1;

  setCell_(existing, headerMap, "Timestamp", existing[headerMap.Timestamp - 1] || now);
  setCell_(existing, headerMap, "Name", data.name);
  setCell_(existing, headerMap, "Email", data.email);
  setCell_(existing, headerMap, "Phone", data.phone || "");
  setCell_(existing, headerMap, "Status", data.status);
  setCell_(existing, headerMap, "Adults", data.status === "yes" ? data.adults || "0" : "0");
  setCell_(existing, headerMap, "Children", data.status === "yes" ? data.children || "0" : "0");
  setCell_(existing, headerMap, "Message", data.message || "");
  setCell_(existing, headerMap, "Submitted At", data.date || now);
  setCell_(existing, headerMap, "Confirmation Sent At", "");

  if (data.status !== "yes") {
    setCell_(existing, headerMap, "Reminder Sent At", "");
  }

  sheet.getRange(rowNumber, 1, 1, existing.length).setValues([existing]);
  return rowNumber;
}

function findRowsByEmail_(sheet, headerMap, email) {
  const emailColumn = headerMap.Email;
  const lastRow = sheet.getLastRow();
  if (!emailColumn || lastRow < 2) return [];

  const emails = sheet.getRange(2, emailColumn, lastRow - 1, 1).getValues();
  return emails.reduce((rows, row, index) => {
    if (String(row[0] || "").trim().toLowerCase() === email) rows.push(index + 2);
    return rows;
  }, []);
}

function setCell_(row, headerMap, header, value) {
  row[headerMap[header] - 1] = value;
}

function sendConfirmationEmail_(data) {
  const attending = data.status === "yes";
  const subject = attending
    ? "RSVP confirmed: " + CONFIG.eventTitle
    : "RSVP received: " + CONFIG.eventTitle;
  const greeting = "Dear " + escapeHtml_(data.name) + ",";
  const intro = attending
    ? "Thank you for your RSVP. We are delighted that you can join us for this special celebration."
    : "Thank you for letting us know that you cannot make it. We will miss celebrating with you.";
  const calendarBlock = attending ? calendarButtonsHtml_() : "";
  const guestLine = attending ? "<p><strong>Guests:</strong> " + escapeHtml_(formatGuestCount_(data)) + "</p>" : "";
  const htmlBody = emailShell_(greeting, intro + guestLine + eventDetailsHtml_() + calendarBlock);
  const plainBody = attending
    ? "Thank you for your RSVP.\n\n" + eventDetailsText_() + "\n\nAdd to Google Calendar: " + CONFIG.googleCalendarUrl + "\nCalendar invite: " + CONFIG.icsUrl
    : "Thank you for letting us know that you cannot make it.\n\n" + eventDetailsText_();

  sendEmail_(data.email, subject, plainBody, htmlBody, attending);
}

function sendReminderEmail_(data) {
  const subject = "Reminder: " + CONFIG.eventTitle + " is coming up";
  const greeting = "Dear " + escapeHtml_(data.name || "Guest") + ",";
  const intro = "We are looking forward to seeing you soon. Here are the event details again for your calendar.";
  const guestLine = "<p><strong>Guests:</strong> " + escapeHtml_(formatGuestCount_(data)) + "</p>";
  const htmlBody = emailShell_(greeting, intro + guestLine + eventDetailsHtml_() + calendarButtonsHtml_());
  const plainBody = "We are looking forward to seeing you soon.\n\n" + eventDetailsText_() + "\n\nAdd to Google Calendar: " + CONFIG.googleCalendarUrl + "\nCalendar invite: " + CONFIG.icsUrl;

  sendEmail_(data.email, subject, plainBody, htmlBody, true);
}

function sendEmail_(to, subject, body, htmlBody, attachInvite) {
  const options = {
    name: "Aadhya & Thanvi RSVP",
    htmlBody
  };

  if (attachInvite) {
    options.attachments = [
      Utilities.newBlob(buildIcs_(), "text/calendar", "aadhya-thanvi-ranga-pravesham.ics")
    ];
  }

  MailApp.sendEmail(to, subject, body, options);
}

function eventDetailsHtml_() {
  return [
    "<div style=\"margin:22px 0;padding:18px;border:1px solid #ead8aa;border-radius:10px;background:#fffaf0;\">",
    "<p style=\"margin:0 0 8px;\"><strong>Date:</strong> " + escapeHtml_(CONFIG.eventDateLabel) + "</p>",
    "<p style=\"margin:0 0 8px;\"><strong>Time:</strong> " + escapeHtml_(CONFIG.eventTimeLabel) + "</p>",
    "<p style=\"margin:0 0 8px;\"><strong>Doors:</strong> " + escapeHtml_(CONFIG.doorsLabel) + "</p>",
    "<p style=\"margin:0;\"><strong>Venue:</strong> " + escapeHtml_(CONFIG.eventLocation) + "</p>",
    "</div>"
  ].join("");
}

function eventDetailsText_() {
  return [
    CONFIG.eventTitle,
    "Date: " + CONFIG.eventDateLabel,
    "Time: " + CONFIG.eventTimeLabel,
    "Doors: " + CONFIG.doorsLabel,
    "Venue: " + CONFIG.eventLocation,
    "Website: " + CONFIG.siteUrl
  ].join("\n");
}

function calendarButtonsHtml_() {
  return [
    "<p style=\"margin:22px 0 10px;\">Add the event to your calendar:</p>",
    "<p>",
    "<a href=\"" + CONFIG.googleCalendarUrl + "\" style=\"display:inline-block;margin:0 8px 10px 0;padding:11px 16px;border-radius:999px;background:#c9a84c;color:#2a1309;text-decoration:none;font-weight:bold;\">Add to Google Calendar</a>",
    "<a href=\"" + CONFIG.icsUrl + "\" style=\"display:inline-block;margin:0 0 10px;padding:11px 16px;border-radius:999px;border:1px solid #c9a84c;color:#5a1525;text-decoration:none;font-weight:bold;\">Download Calendar Invite</a>",
    "</p>"
  ].join("");
}

function emailShell_(greeting, bodyHtml) {
  return [
    "<div style=\"font-family:Arial,sans-serif;line-height:1.6;color:#2c1a0e;max-width:620px;\">",
    "<h2 style=\"font-family:Georgia,serif;color:#5a1525;margin-bottom:8px;\">Aadhya &amp; Thanvi's Ranga Pravesham</h2>",
    "<p>" + greeting + "</p>",
    "<div>" + bodyHtml + "</div>",
    "<p style=\"margin-top:24px;\">With gratitude,<br>The Sure, Pasam, Nagineni &amp; Pachava families</p>",
    "</div>"
  ].join("");
}

function formatGuestCount_(data) {
  const adults = Number(data.adults || 0);
  const children = Number(data.children || 0);
  const parts = [];
  if (adults) parts.push(adults + " adult" + (adults === 1 ? "" : "s"));
  if (children) parts.push(children + " child" + (children === 1 ? "" : "ren"));
  return parts.length ? parts.join(", ") : "0 guests";
}

function normaliseStatus_(status) {
  const value = String(status || "").trim().toLowerCase();
  if (["yes", "y", "attending", "attending!"].includes(value)) return "yes";
  if (["no", "n", "declined", "can't make it", "cant make it"].includes(value)) return "no";
  return value;
}

function publicDisplayName_(name) {
  return cleanPublicText_(name, 80) || "Guest";
}

function cleanPublicText_(value, maxLength) {
  const text = String(value || "").replace(/\s+/g, " ").trim();
  if (!text) return "";
  return text.length > maxLength ? text.slice(0, maxLength - 1).trim() + "..." : text;
}

function publicDate_(row, headerMap) {
  const submitted = row[headerMap["Submitted At"] - 1];
  const timestamp = row[headerMap.Timestamp - 1];
  const value = submitted || timestamp;
  const date = value instanceof Date ? value : new Date(value);
  return isNaN(date.getTime()) ? "" : date.toISOString();
}

function buildIcs_() {
  const lines = [
    "BEGIN:VCALENDAR",
    "VERSION:2.0",
    "PRODID:-//Aadhya and Thanvi Ranga Pravesham//EN",
    "CALSCALE:GREGORIAN",
    "METHOD:PUBLISH",
    "BEGIN:VEVENT",
    "UID:aadhya-thanvi-ranga-pravesham-20260725@aadhyathanviarangetram.com",
    "DTSTAMP:20260612T000000Z",
    "DTSTART:" + CONFIG.eventStartUtc,
    "DTEND:" + CONFIG.eventEndUtc,
    "SUMMARY:" + escapeIcs_(CONFIG.eventTitle),
    "LOCATION:" + escapeIcs_(CONFIG.eventLocation),
    "DESCRIPTION:" + escapeIcs_(CONFIG.doorsLabel + ". You are invited to Aadhya Sure and Thanvi Vashishta Nagineni's Bharatanatyam Ranga Pravesham."),
    "URL:" + CONFIG.siteUrl,
    "END:VEVENT",
    "END:VCALENDAR"
  ];

  return lines.map(foldIcsLine_).join("\r\n") + "\r\n";
}

function foldIcsLine_(line) {
  const chunks = [];
  let remaining = line;
  while (remaining.length > 73) {
    chunks.push(remaining.slice(0, 73));
    remaining = " " + remaining.slice(73);
  }
  chunks.push(remaining);
  return chunks.join("\r\n");
}

function escapeIcs_(value) {
  return String(value || "")
    .replace(/\\/g, "\\\\")
    .replace(/;/g, "\\;")
    .replace(/,/g, "\\,")
    .replace(/\n/g, "\\n");
}

function escapeHtml_(value) {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function json_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function jsonp_(callback, payload) {
  if (!/^[A-Za-z_$][0-9A-Za-z_$]*(\.[A-Za-z_$][0-9A-Za-z_$]*)*$/.test(String(callback || ""))) {
    return json_({ ok: false, error: "Invalid callback." });
  }

  return ContentService
    .createTextOutput(callback + "(" + JSON.stringify(payload) + ");")
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}
