/**
 * Google Apps Script — API Cost Logger for UQ Slide Converter (p2000)
 *
 * SETUP:
 * 1. Create a new Google Sheet (name it e.g. "p2000 API Costs")
 * 2. Add these headers in row 1 of Sheet1:
 *      A: timestamp | B: model | C: purpose | D: slide_info | E: filename
 *      F: input_tokens | G: output_tokens | H: total_tokens
 *      I: input_cost_usd | J: output_cost_usd | K: total_cost_usd
 * 3. Open Extensions > Apps Script
 * 4. Paste this entire file into Code.gs (replace any existing code)
 * 5. Click Deploy > New deployment
 *    - Type: Web app
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 6. Copy the web app URL
 * 7. Set it as GOOGLE_SHEETS_WEBHOOK_URL in your Render environment variables
 *
 * The script accepts POST requests with JSON cost data and appends rows.
 * It also accepts GET requests to return all logged data as JSON
 * (used by the admin panel to read persistent cost history).
 */

// --- Configuration ---
const SHEET_NAME = "Sheet1";

// --- POST handler: append a cost log entry ---
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    // Support both single entries and batches
    const entries = Array.isArray(data) ? data : [data];

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sheet) {
      return ContentService
        .createTextOutput(JSON.stringify({ error: "Sheet not found" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const rows = entries.map(entry => [
      entry.timestamp || new Date().toISOString(),
      entry.model || "",
      entry.purpose || "",
      entry.slide_info || "",
      entry.filename || "",
      entry.input_tokens || 0,
      entry.output_tokens || 0,
      entry.total_tokens || 0,
      entry.input_cost_usd || 0,
      entry.output_cost_usd || 0,
      entry.total_cost_usd || 0,
    ]);

    // Batch append for efficiency
    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, rows.length, 11).setValues(rows);

    return ContentService
      .createTextOutput(JSON.stringify({ status: "ok", rows_added: rows.length }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// --- GET handler: return all cost log entries as JSON ---
function doGet(e) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sheet) {
      return ContentService
        .createTextOutput(JSON.stringify({ error: "Sheet not found" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      // Only headers, no data
      return ContentService
        .createTextOutput(JSON.stringify([]))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const data = sheet.getRange(2, 1, lastRow - 1, 11).getValues();

    const entries = data.map(row => ({
      timestamp: row[0],
      model: row[1],
      purpose: row[2],
      slide_info: row[3],
      filename: row[4],
      input_tokens: row[5],
      output_tokens: row[6],
      total_tokens: row[7],
      input_cost_usd: row[8],
      output_cost_usd: row[9],
      total_cost_usd: row[10],
    }));

    return ContentService
      .createTextOutput(JSON.stringify(entries))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
