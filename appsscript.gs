// ─────────────────────────────────────────────────────────────
// Glow-Up Checklist — Google Apps Script
// Paste this entire file into your Apps Script editor,
// then deploy as a Web App (instructions below).
// ─────────────────────────────────────────────────────────────

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // ── Summary sheet ──────────────────────────────────────
    const summarySheet = getOrCreateSheet(ss, 'Daily Summary', [
      'Date', 'Day', 'Daily Score', 'Daily %',
      'Weekly Items Done', 'Monthly Items Done', 'Last Synced'
    ]);

    const sumRow = [
      data.summary.date,
      data.summary.day,
      data.summary.score,
      data.summary.percentage + '%',
      data.summary.weeklyDone,
      data.summary.monthlyDone,
      new Date().toLocaleString('en-AU')
    ];
    upsertRow(summarySheet, data.summary.date, sumRow, 0);

    // ── Detail sheet ───────────────────────────────────────
    const detailSheet = getOrCreateSheet(ss, 'Item Detail', [
      'Date', 'Tab', 'Section', 'Item', 'Checked', 'Last Synced'
    ]);

    data.details.forEach(function(item) {
      const row = [
        data.summary.date,
        item.tab,
        item.section,
        item.label,
        item.checked ? 'YES' : 'no',
        new Date().toLocaleString('en-AU')
      ];
      // Match on date + item label
      upsertRow(detailSheet, data.summary.date + '|' + item.label, row, null, item.label);
    });

    return jsonResponse({ ok: true });
  } catch (err) {
    return jsonResponse({ ok: false, error: err.message });
  }
}

function doGet() {
  return jsonResponse({ ok: true, status: 'Glow-Up API is running ✨' });
}

// ── Helpers ────────────────────────────────────────────────

function getOrCreateSheet(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function upsertRow(sheet, matchKey, row, colA, colB) {
  // Find existing row matching on col A (and optionally col B = item label)
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const key = colB != null
      ? data[i][0] + '|' + data[i][3]   // date + item label
      : String(data[i][0]);              // date only
    if (key === matchKey) {
      sheet.getRange(i + 1, 1, 1, row.length).setValues([row]);
      return;
    }
  }
  sheet.appendRow(row);
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
