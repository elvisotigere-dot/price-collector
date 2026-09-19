// ═══════════════════════════════════════════════════════
//  PRICE COLLECTOR — Google Apps Script
//  Paste this entire file into your Apps Script editor
// ═══════════════════════════════════════════════════════

const SHEET_NAME = 'Supplier Prices';
const DATA_START_ROW = 5;    // Your items start at row 5
const ITEM_COL = 1;          // Column A = items
const SUPPLIER_START_COL = 3; // Column C = first supplier (Bilal)

function doGet(e) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    const data = sheet.getRange(1, 1, lastRow, lastCol).getValues();

    // Read supplier names from row 3 (index 2), starting at column C (index 2)
    const suppliers = [];
    for (let c = SUPPLIER_START_COL - 1; c < lastCol; c++) {
      const name = String(data[2][c] || '').trim();
      if (name) suppliers.push({ name: name, col: c + 1 });
    }

    // Read items from row 5 onwards
    const items = [];
    for (let r = DATA_START_ROW - 1; r < lastRow; r++) {
      const itemName = String(data[r][ITEM_COL - 1] || '').trim();
      if (!itemName) continue;

      const prices = {};
      suppliers.forEach(s => {
        const val = data[r][s.col - 1];
        prices[s.name] = (val === '' || val === null || val === undefined) ? '' : val;
      });

      items.push({
        name: itemName,
        row: r + 1,
        prices: prices
      });
    }

    const result = {
      suppliers: suppliers.map(s => s.name),
      items: items
    };

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    const lastCol = sheet.getLastColumn();

    // Build supplier → column map from row 3
    const supplierRow = sheet.getRange(3, 1, 1, lastCol).getValues()[0];
    const supplierCols = {};
    for (let c = SUPPLIER_START_COL - 1; c < supplierRow.length; c++) {
      const name = String(supplierRow[c] || '').trim();
      if (name) supplierCols[name] = c + 1;
    }

    const payload = JSON.parse(e.postData.contents);
    const prices = payload.prices || [];

    prices.forEach(p => {
      const col = supplierCols[p.supplier];
      if (col && p.row) {
        sheet.getRange(p.row, col).setValue(p.price);
      }
    });

    SpreadsheetApp.flush();

    return ContentService
      .createTextOutput(JSON.stringify({ success: true, saved: prices.length }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
